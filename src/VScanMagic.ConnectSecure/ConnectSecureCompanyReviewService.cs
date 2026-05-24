using System.Text.Json;
using System.Text.RegularExpressions;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureCompanyReviewService(ConnectSecureClient client)
{
    private static readonly Regex IpTargetRegex = new(@"^\d+\.\d+\.\d+\.\d+(/(\d+))?$", RegexOptions.Compiled);

    public async Task<CompanyReviewData> GetReviewDataAsync(int companyId, string companyName = "", CancellationToken ct = default)
    {
        var result = new CompanyReviewData { CompanyId = companyId, CompanyName = companyName };
        if (companyId <= 0)
            return result;

        var lwAssets = await FetchArrayAsync("/r/report_queries/lightweight_assets",
            LightweightAssetsQuery(companyId), ct);
        result.AgentCount = lwAssets.Count;

        var agents = await FetchArrayAsync("/r/company/agents", CompanyQuery(companyId), ct);
        if (result.AgentCount == 0 && agents.Count > 0)
        {
            var lightweight = agents.Count(a =>
                ConnectSecureJsonReader.GetString(a, "agent_type", "agentType")
                    .Contains("lightweight", StringComparison.OrdinalIgnoreCase));
            result.AgentCount = lightweight > 0 ? lightweight : agents.Count;
        }

        var credMappings = await FetchArrayAsync("/r/company/agent_credentials_mapping", CompanyQuery(companyId), ct);
        var discMappings = await FetchArrayAsync("/r/company/agent_discoverysettings_mapping", CompanyQuery(companyId), ct);

        var agentIdsWithCreds = credMappings
            .Select(m => ConnectSecureJsonReader.GetInt(m, "agent_id", "agentId"))
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToHashSet();
        var agentIdsWithNetworks = discMappings
            .Select(m => ConnectSecureJsonReader.GetInt(m, "agent_id", "agentId"))
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToHashSet();

        result.ProbesWithCredentials = agentIdsWithCreds.Count;
        result.ProbesWithNetworks = agentIdsWithNetworks.Count;
        result.ProbesWithBoth = agentIdsWithCreds.Count(id => agentIdsWithNetworks.Contains(id));

        await PopulateProbeNmapInfoAsync(result, agents, companyId, ct);
        await PopulateDiscoverySettingsAsync(result, discMappings, companyId, ct);
        PopulateOfflineAgents(result, agents);

        await PopulateFirewallAsync(result, companyId, ct);
        await PopulateScanDatesAsync(result, agents, agentIdsWithCreds, agentIdsWithNetworks, companyId, ct);
        PopulateQuickWins(result);

        return result;
    }

    public static IReadOnlyList<CompanyReviewCheck> BuildChecks(CompanyReviewData data) =>
    [
        new("1. Lightweight agents", data.AgentCount.ToString(), data.AgentCount > 0),
        new("2. Probes w/ creds + networks", data.ProbesWithBoth.ToString(), data.ProbesWithBoth > 0),
        new("3. External scan targets", data.ExternalAssets.Count.ToString(), data.ExternalAssets.Count > 0),
        new("4. Offline (7d / 14d / 30+d)",
            $"{data.AgentsOffline7PlusDays} / {data.AgentsOffline14PlusDays} / {data.AgentsOffline30PlusDays}",
            data.AgentsOffline30PlusDays == 0),
        new("5. Firewall integration",
            data.FirewallActive
                ? $"{data.FirewallCount} firewall(s): {data.FirewallType}"
                : "Not configured",
            data.FirewallActive),
        new("6. Last internal scan", data.LastInternalScan ?? "None", !string.IsNullOrWhiteSpace(data.LastInternalScan)),
        new("7. Last external scan", data.LastExternalScan ?? "None", !string.IsNullOrWhiteSpace(data.LastExternalScan))
    ];

    public static IReadOnlyList<string> CombineSubnetLines(CompanyReviewData data)
    {
        var lines = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var subnet in data.ProbesSubnets)
            lines.Add(subnet);
        foreach (var target in data.ScanTargets)
            lines.Add(target);
        return lines.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private async Task PopulateProbeNmapInfoAsync(
        CompanyReviewData result,
        IReadOnlyList<JsonElement> agents,
        int companyId,
        CancellationToken ct)
    {
        var probeAgents = await FetchArrayAsync("/r/company/agent_discovery_credentials",
            ProbeAgentsQuery(companyId), ct);

        if (probeAgents.Count == 0)
        {
            probeAgents = agents.Where(a =>
            {
                if (a.TryGetProperty("probe_setting", out var ps) || a.TryGetProperty("probeSetting", out ps))
                    return ps.ValueKind == JsonValueKind.Object;
                return false;
            }).ToList();
        }

        foreach (var agent in probeAgents)
        {
            var agentId = ConnectSecureJsonReader.GetInt(agent, "id");
            var hostName = FirstNonEmpty(
                ConnectSecureJsonReader.GetString(agent, "host_name", "hostName", "Host Name"),
                ConnectSecureJsonReader.GetString(agent, "agent_name", "agentName", "hostname", "name"),
                "(unnamed)");
            var ip = FirstNonEmpty(
                ConnectSecureJsonReader.GetString(agent, "ip", "agent_ip"),
                "(none)");
            var nmap = ConnectSecureJsonReader.GetString(agent, "nmap_interface", "nmapInterface");
            if (string.IsNullOrWhiteSpace(nmap))
                nmap = "(not set)";

            var interfaces = ProbeInterfaceHelper.ParseAvailableInterfaces(agent);

            string? port = null;
            if (agent.TryGetProperty("probe_setting", out var probeSetting) ||
                agent.TryGetProperty("probeSetting", out probeSetting))
            {
                port = FirstNonEmpty(
                    ConnectSecureJsonReader.GetString(probeSetting, "listen_port", "listenPort"),
                    ConnectSecureJsonReader.GetString(probeSetting, "port", "nmap_port"));
            }

            result.ProbeAgentsNmapInfo.Add(new ProbeNmapInfo(agentId, hostName, ip, nmap, port, interfaces));
        }
    }

    private async Task PopulateDiscoverySettingsAsync(
        CompanyReviewData result,
        IReadOnlyList<JsonElement> discMappings,
        int companyId,
        CancellationToken ct)
    {
        var discoverySettings = await FetchDiscoverySettingsAsync(companyId, ct);
        var dsIdToAddr = new Dictionary<int, string>();

        foreach (var ds in discoverySettings)
        {
            var dsId = ConnectSecureJsonReader.GetInt(ds, "id");
            var address = ConnectSecureJsonReader.GetString(ds, "address");
            if (dsId.HasValue && !string.IsNullOrWhiteSpace(address))
                dsIdToAddr[dsId.Value] = address;

            AddDiscoveryTargets(result, ds, address);

            if (IpTargetRegex.IsMatch(address.Trim()))
            {
                var targetsForValidation = new List<string> { address.Trim() };
                AppendTargetValues(targetsForValidation, ds);
                foreach (var issue in ExternalSubnetHelper.ValidateExternalTargets(targetsForValidation, address.Trim()))
                {
                    if (!result.SubnetIssues.Contains(issue))
                        result.SubnetIssues.Add(issue);
                }
            }
        }

        foreach (var mapping in discMappings)
        {
            PopulateInternalSubnetFromMapping(result, mapping, discoverySettings, dsIdToAddr);
        }

        var extScan = discoverySettings.Where(IsExternalDiscoverySetting).ToList();
        foreach (var ds in extScan)
        {
            var dsId = ConnectSecureJsonReader.GetInt(ds, "id");
            var name = ConnectSecureJsonReader.GetString(ds, "discovery_settings_name", "name");
            var address = ConnectSecureJsonReader.GetString(ds, "address");
            var targetIp = ConnectSecureJsonReader.GetString(ds, "target_ip", "targetIp", "target_ips");
            var addr = FirstNonEmpty(address, targetIp);
            if (string.IsNullOrWhiteSpace(addr))
                continue;

            var scanIps = ResolveExternalScanIps(address, targetIp);
            result.ExternalAssets.Add(new ExternalAssetEntry(
                dsId,
                string.IsNullOrWhiteSpace(name) ? "(unnamed)" : name,
                addr,
                targetIp,
                scanIps.Count));

            foreach (var part in scanIps)
            {
                if (IpTargetRegex.IsMatch(part) && !result.ScanTargets.Contains(part))
                    result.ScanTargets.Add(part);
            }

            if (ExternalSubnetHelper.GetSubnetBounds(address) is not null)
            {
                foreach (var issue in ExternalSubnetHelper.ValidateExternalTargets(scanIps, address))
                {
                    if (!result.SubnetIssues.Contains(issue))
                        result.SubnetIssues.Add(issue);
                }
            }
        }
    }

    private async Task<List<JsonElement>> FetchDiscoverySettingsAsync(int companyId, CancellationToken ct)
    {
        var fromReport = await FetchArrayAsync("/r/report_queries/discovery_settings",
            CompanyQuery(companyId, limit: 500, orderBy: "updated desc"), ct);
        if (fromReport.Count > 0)
            return fromReport;

        var all = await FetchArrayAsync("/r/company/discovery_settings",
            new Dictionary<string, string> { ["limit"] = "2000", ["skip"] = "0" }, ct);
        return all
            .Where(ds => ConnectSecureJsonReader.GetInt(ds, "company_id", "companyId") == companyId)
            .ToList();
    }

    private static void AddDiscoveryTargets(CompanyReviewData result, JsonElement ds, string address)
    {
        var targets = new List<string>();
        if (!string.IsNullOrWhiteSpace(address))
        {
            foreach (var part in address.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (IpTargetRegex.IsMatch(part))
                    targets.Add(part);
            }
        }

        AppendTargetValues(targets, ds);
        foreach (var target in targets)
        {
            if (!result.ScanTargets.Contains(target))
                result.ScanTargets.Add(target);
        }
    }

    private static void AppendTargetValues(List<string> targets, JsonElement ds)
    {
        var tip = ConnectSecureJsonReader.GetString(ds, "target_ip", "target_ips", "targetIp");
        if (string.IsNullOrWhiteSpace(tip))
            return;

        foreach (var part in tip.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!string.IsNullOrWhiteSpace(part))
                targets.Add(part);
        }
    }

    public static bool IsExternalDiscoverySetting(JsonElement ds)
    {
        var type = ConnectSecureJsonReader.GetString(ds, "discovery_settings_type", "type");
        return type.Contains("external", StringComparison.OrdinalIgnoreCase);
    }

    private static void PopulateInternalSubnetFromMapping(
        CompanyReviewData result,
        JsonElement mapping,
        IReadOnlyList<JsonElement> discoverySettings,
        IReadOnlyDictionary<int, string> dsIdToAddr)
    {
        var dsId = ConnectSecureJsonReader.GetInt(mapping, "discovery_settings_id", "discoverysettings_id");
        if (dsId is null)
            return;

        var ds = discoverySettings.FirstOrDefault(item => ConnectSecureJsonReader.GetInt(item, "id") == dsId);
        if (ds.ValueKind == JsonValueKind.Undefined || IsExternalDiscoverySetting(ds))
            return;

        var address = FirstNonEmpty(
            dsIdToAddr.GetValueOrDefault(dsId.Value),
            ConnectSecureJsonReader.GetString(ds, "address"),
            ConnectSecureJsonReader.GetString(ds, "target_ip", "targetIp", "target_ips"));
        if (string.IsNullOrWhiteSpace(address))
            return;

        var mappingId = ConnectSecureJsonReader.GetInt(mapping, "id");
        var agentId = ConnectSecureJsonReader.GetInt(mapping, "agent_id", "agentId");
        var name = FirstNonEmpty(
            ConnectSecureJsonReader.GetString(ds, "discovery_settings_name", "name"),
            address);
        var targetIp = ConnectSecureJsonReader.GetString(ds, "target_ip", "targetIp", "target_ips");
        var scanIps = ResolveExternalScanIps(address, targetIp);
        var probeHostName = result.ProbeAgentsNmapInfo
            .Where(p => p.AgentId == agentId)
            .Select(p => p.HostName)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(probeHostName))
            probeHostName = agentId is > 0 ? $"Agent #{agentId}" : "(unmapped)";

        result.InternalSubnets.Add(new InternalSubnetEntry(
            dsId,
            mappingId,
            agentId,
            probeHostName,
            string.IsNullOrWhiteSpace(name) ? "(unnamed)" : name,
            address,
            targetIp,
            scanIps.Count));

        if (!result.ProbesSubnets.Contains(address))
            result.ProbesSubnets.Add(address);

        if (ExternalSubnetHelper.GetSubnetBounds(address) is not null)
        {
            foreach (var issue in ExternalSubnetHelper.ValidateExternalTargets(scanIps, address))
            {
                if (!result.SubnetIssues.Contains(issue))
                    result.SubnetIssues.Add(issue);
            }
        }
    }

    private static bool IsExternalScanSetting(JsonElement ds) => IsExternalDiscoverySetting(ds);

    private static void PopulateOfflineAgents(CompanyReviewData result, IReadOnlyList<JsonElement> agents)
    {
        var now = DateTime.UtcNow;
        foreach (var agent in agents)
        {
            if (ConnectSecureJsonReader.GetBool(agent, "is_deprecated", "isDeprecated"))
                continue;

            var lastPing = ConnectSecureJsonReader.GetString(agent, "last_ping_time", "lastPingTime", "Last Ping Time");
            if (string.IsNullOrWhiteSpace(lastPing) || !DateTime.TryParse(lastPing, out var pingTime))
                continue;

            var days = (now - pingTime.ToUniversalTime()).TotalDays;
            if (days >= 7) result.AgentsOffline7PlusDays++;
            if (days >= 14) result.AgentsOffline14PlusDays++;
            if (days >= 30)
            {
                result.AgentsOffline30PlusDays++;
                var name = ResolveAgentName(agent);
                result.AgentsOffline30PlusNames.Add(name);
            }
        }
    }

    private static string ResolveAgentName(JsonElement agent)
    {
        var name = FirstNonEmpty(
            ConnectSecureJsonReader.GetString(agent, "host_name", "hostName", "Host Name"),
            ConnectSecureJsonReader.GetString(agent, "agent_name", "agentName", "hostname", "name", "computer_name", "display_name", "asset_name"));
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        var ip = ConnectSecureJsonReader.GetString(agent, "ip");
        if (!string.IsNullOrWhiteSpace(ip))
            return $"IP: {ip}";

        var id = ConnectSecureJsonReader.GetInt(agent, "id");
        return id.HasValue ? $"Agent #{id.Value}" : "Unknown";
    }

    private async Task PopulateFirewallAsync(CompanyReviewData result, int companyId, CancellationToken ct)
    {
        var fwAssets = await FetchArrayAsync("/r/report_queries/firewall_asset_view",
            new Dictionary<string, string>
            {
                ["condition"] = $"company_id={companyId}",
                ["limit"] = "500",
                ["skip"] = "0",
                ["order_by"] = "host_name asc"
            }, ct);

        result.FirewallCount = fwAssets.Count;
        var manufacturers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var asset in fwAssets)
        {
            if (!ConnectSecureJsonReader.GetBool(asset, "is_firewall", "isFirewall"))
                continue;

            var mfr = ConnectSecureJsonReader.GetString(asset, "manufacturer", "asset_type");
            if (!string.IsNullOrWhiteSpace(mfr))
                manufacturers.Add(mfr);
        }

        if (manufacturers.Count == 0 && result.FirewallCount > 0)
            manufacturers.Add("Managed firewall");

        result.FirewallActive = result.FirewallCount > 0;
        result.FirewallType = string.Join(", ", manufacturers.OrderBy(x => x));
    }

    private async Task PopulateScanDatesAsync(
        CompanyReviewData result,
        IReadOnlyList<JsonElement> agents,
        HashSet<int> agentIdsWithCreds,
        HashSet<int> agentIdsWithNetworks,
        int companyId,
        CancellationToken ct)
    {
        var probeIds = agentIdsWithCreds.Where(id => agentIdsWithNetworks.Contains(id)).ToHashSet();
        DateTime? lastInternal = null;
        foreach (var agent in agents)
        {
            var id = ConnectSecureJsonReader.GetInt(agent, "id");
            if (!id.HasValue || !probeIds.Contains(id.Value))
                continue;

            var lastScanned = ConnectSecureJsonReader.GetString(agent, "last_scanned_time", "lastScannedTime");
            if (string.IsNullOrWhiteSpace(lastScanned) || !DateTime.TryParse(lastScanned, out var scanned))
                continue;

            if (lastInternal is null || scanned > lastInternal)
                lastInternal = scanned;
        }

        if (lastInternal.HasValue)
            result.LastInternalScan = lastInternal.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");

        if (string.IsNullOrWhiteSpace(result.LastInternalScan))
        {
            var stats = await FetchCompanyStatsAsync(companyId, ct);
            if (stats.HasValue)
            {
                result.LastInternalScan = FirstNonEmpty(
                    ConnectSecureJsonReader.GetString(stats.Value, "ad_last_scan_time", "date"));
            }
        }

        DateTime? lastExternal = null;
        var jobs = await FetchArrayAsync("/r/company/jobs_view",
            new Dictionary<string, string>
            {
                ["condition"] = $"company_id={companyId} and type='External Scan'",
                ["limit"] = "50",
                ["skip"] = "0",
                ["order_by"] = "updated desc"
            }, ct);

        foreach (var job in jobs)
        {
            var updated = ConnectSecureJsonReader.GetString(job, "updated", "created");
            if (string.IsNullOrWhiteSpace(updated) || !DateTime.TryParse(updated, out var dt))
                continue;
            if (lastExternal is null || dt > lastExternal)
                lastExternal = dt;
        }

        if (lastExternal.HasValue)
            result.LastExternalScan = lastExternal.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");

        if (string.IsNullOrWhiteSpace(result.LastExternalScan))
        {
            var stats = await FetchCompanyStatsAsync(companyId, ct);
            if (stats.HasValue)
            {
                var ext = FirstNonEmpty(
                    ConnectSecureJsonReader.GetString(stats.Value, "external_last_scan_time", "date", "updated"));
                if (!string.IsNullOrWhiteSpace(ext) && DateTime.TryParse(ext, out var parsed))
                    result.LastExternalScan = parsed.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
                else
                    result.LastExternalScan = ext;
            }
        }
    }

    private static void PopulateQuickWins(CompanyReviewData result)
    {
        if (result.AgentCount == 0)
            result.QuickWins.Add("Add lightweight agents to enable internal scanning");
        if (result.ProbesWithBoth == 0)
            result.QuickWins.Add("Map credentials and discovery networks to at least one probe agent");
        var probesWithNmap = result.ProbeAgentsNmapInfo.Count(p =>
            !string.IsNullOrWhiteSpace(p.NmapInterface) && p.NmapInterface != "(not set)");
        if (result.ProbesWithBoth > 0 && probesWithNmap == 0)
            result.QuickWins.Add("Configure nmap interface on probe agent(s) for scanning");
        if (result.SubnetIssues.Count > 0)
            result.QuickWins.Add("Exclude network, ISP gateway, and broadcast addresses from scan targets");
        if (result.AgentsOffline30PlusDays > 0)
            result.QuickWins.Add("Investigate agents/probes offline more than 30 days; reinstall or remove");
        if (!result.FirewallActive)
            result.QuickWins.Add("Configure firewall integration for visibility");
        if (string.IsNullOrWhiteSpace(result.LastInternalScan) && string.IsNullOrWhiteSpace(result.LastExternalScan))
            result.QuickWins.Add("Run internal and external scans to populate vulnerability data");
    }

    public static void RebuildQuickWins(CompanyReviewData result)
    {
        result.QuickWins.Clear();
        PopulateQuickWins(result);
    }

    private async Task<List<JsonElement>> FetchArrayAsync(
        string endpoint,
        IReadOnlyDictionary<string, string> query,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, endpoint, query, ct: ct);
        return ConnectSecureJsonReader.ExtractDataArray(response);
    }

    private async Task<JsonElement?> FetchCompanyStatsAsync(int companyId, CancellationToken ct)
    {
        var rows = await FetchArrayAsync("/r/company/company_stats", CompanyQuery(companyId, limit: 100), ct);
        if (rows.Count == 0)
            return null;

        JsonElement? best = null;
        DateTime? bestDate = null;
        foreach (var row in rows)
        {
            var dateText = FirstNonEmpty(
                ConnectSecureJsonReader.GetString(row, "date"),
                ConnectSecureJsonReader.GetString(row, "ad_last_scan_time", "adLastScanTime"));
            if (!DateTime.TryParse(dateText, out var date))
                continue;
            if (bestDate is null || date > bestDate)
            {
                bestDate = date;
                best = row;
            }
        }

        return best ?? rows[0];
    }

    internal static Dictionary<string, string> CompanyQuery(int companyId, int limit = 5000, int skip = 0, string? orderBy = null)
    {
        var query = new Dictionary<string, string>
        {
            ["condition"] = $"company_id={companyId}",
            ["limit"] = limit.ToString(),
            ["skip"] = skip.ToString()
        };
        if (!string.IsNullOrWhiteSpace(orderBy))
            query["order_by"] = orderBy;
        return query;
    }

    internal static Dictionary<string, string> LightweightAssetsQuery(int companyId) =>
        new()
        {
            ["company_id"] = companyId.ToString(),
            ["limit"] = "5000",
            ["skip"] = "0"
        };

    internal static Dictionary<string, string> ProbeAgentsQuery(int companyId) =>
        new()
        {
            ["condition"] = $"company_id={companyId} and agent_type='PROBE' and is_deprecated=FALSE and is_retired=FALSE",
            ["skip"] = "0",
            ["limit"] = "100",
            ["order_by"] = "host_name asc"
        };


    private static List<string> ResolveExternalScanIps(string? address, string? targetIp)
    {
        if (!string.IsNullOrWhiteSpace(targetIp))
        {
            return targetIp
                .Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(p => IpTargetRegex.IsMatch(p) && !p.Contains('/'))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (!string.IsNullOrWhiteSpace(address) && ExternalSubnetHelper.GetSubnetBounds(address) is not null)
            return ExternalSubnetHelper.ExpandCidrToUsableIps(address).ToList();

        if (!string.IsNullOrWhiteSpace(address))
        {
            return address
                .Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(p => IpTargetRegex.IsMatch(p) && !p.Contains('/'))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        return [];
    }

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }

        return "";
    }
}
