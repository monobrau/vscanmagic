using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

namespace VScanMagic.ConnectSecure;

/// <summary>
/// Single source of truth for ConnectSecure remediation data for a company.
///
/// <para>
/// Uses <c>/r/report_queries/get_remediation</c> with
/// <c>condition=company_id=X</c>. This endpoint reliably filters server-side
/// (verified on pod104) and returns one row per (host × fix) with all the
/// metadata needed by the patch and suppress flows: <c>solution_id</c>,
/// <c>product</c>, <c>fix</c>, <c>severity</c>, <c>remediation_action</c>,
/// <c>asset_id</c>, <c>agent_id</c>, <c>host_name</c>, vuln counts, etc.
/// </para>
///
/// <para>
/// Replaces the previous mix of <c>remediation_plan_include_company</c> +
/// <c>remediation_plan_by_company</c> fallbacks, which were unreliable on
/// pod104 (include_company returns empty for some companies, by_company
/// ignores the <c>condition</c> parameter and forces tenant-wide pagination).
/// </para>
/// </summary>
public sealed class ConnectSecureRemediationService(
    ConnectSecureClient client,
    ConnectSecureCacheService cache)
{
    private const string Endpoint = "/r/report_queries/get_remediation";
    private const int PageSize = 5000;
    private const int MaxPages = 10;

    public async Task<RemediationDataset> GetRemediationDatasetAsync(
        int companyId,
        bool forceRefresh = false,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        if (!forceRefresh && cache.TryGetRemediationDataset(companyId, out var cached))
            return cached;

        var sw = Stopwatch.StartNew();
        var records = await FetchAllRecordsAsync(companyId, ct);
        sw.Stop();
        ConnectSecureRequestMetrics.LogPageFetch(
            "remediation_dataset", page: 0, batchRows: records.Count, keptRows: records.Count, elapsedMs: sw.ElapsedMilliseconds);

        var dataset = new RemediationDataset(companyId, DateTimeOffset.UtcNow, records);
        cache.SetRemediationDataset(companyId, dataset);
        return dataset;
    }

    public void Invalidate(int companyId) => cache.InvalidateRemediationDataset(companyId);

    private async Task<List<RemediationRecord>> FetchAllRecordsAsync(int companyId, CancellationToken ct)
    {
        var records = new List<RemediationRecord>();
        var condition = $"company_id={companyId}";

        for (var page = 0; page < MaxPages; page++)
        {
            ct.ThrowIfCancellationRequested();

            var query = new Dictionary<string, string>
            {
                ["condition"] = condition,
                ["limit"] = PageSize.ToString(CultureInfo.InvariantCulture),
                ["skip"] = (page * PageSize).ToString(CultureInfo.InvariantCulture),
                ["order_by"] = "severity desc"
            };

            var sw = Stopwatch.StartNew();
            var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, Endpoint, query, ct: ct);
            var rows = ConnectSecureJsonReader.ExtractDataArray(response);
            sw.Stop();
            ConnectSecureRequestMetrics.LogPageFetch(
                "get_remediation", page, batchRows: rows.Count, keptRows: rows.Count, elapsedMs: sw.ElapsedMilliseconds);

            if (rows.Count == 0)
                break;

            foreach (var row in rows)
            {
                // Defensive: trust the server filter but drop any cross-company leakage.
                var rowCompanyId = ConnectSecureJsonReader.GetInt(row, "company_id", "companyId");
                if (rowCompanyId is > 0 && rowCompanyId.Value != companyId)
                    continue;

                records.Add(ParseRecord(row));
            }

            if (rows.Count < PageSize)
                break;
        }

        return records;
    }

    internal static RemediationRecord ParseRecord(JsonElement row) =>
        new(
            SolutionId: ConnectSecureJsonReader.GetInt(row, "solution_id", "solutionId") ?? 0,
            Product: ConnectSecureJsonReader.GetString(row, "product"),
            Fix: ConnectSecureJsonReader.GetString(row, "fix"),
            FixUrl: ConnectSecureJsonReader.GetString(row, "url"),
            Severity: ConnectSecureJsonReader.GetString(row, "severity"),
            RemediationAction: ConnectSecureJsonReader.GetString(row, "remediation_action", "remediationAction"),
            OsType: ConnectSecureJsonReader.GetString(row, "os_type", "osType"),
            AssetId: ConnectSecureJsonReader.GetInt(row, "asset_id", "assetId", "id") ?? 0,
            AgentId: ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
            HostName: ConnectSecureJsonReader.GetString(row, "host_name", "hostName"),
            Ip: PickFirstIp(row),
            OnlineStatus: ConnectSecureJsonReader.GetBool(row, "online_status", "onlineStatus"),
            TotalVulsCount: ConnectSecureJsonReader.GetInt(row, "total_vuls_count", "totalVulsCount") ?? 0,
            CriticalVulsCount: ConnectSecureJsonReader.GetInt(row, "critical_vuls_count", "criticalVulsCount") ?? 0,
            HighVulsCount: ConnectSecureJsonReader.GetInt(row, "high_vuls_count", "highVulsCount") ?? 0,
            MediumVulsCount: ConnectSecureJsonReader.GetInt(row, "medium_vuls_count", "mediumVulsCount") ?? 0,
            LowVulsCount: ConnectSecureJsonReader.GetInt(row, "low_vuls_count", "lowVulsCount") ?? 0,
            EpssVuls: GetDouble(row, "epss_vuls", "epssVuls"),
            InstallSource: ConnectSecureJsonReader.GetString(row, "install_source", "installSource"),
            FirstVulDiscovered: ParseTimestamp(row, "first_vul_discovered", "firstVulDiscovered"),
            LastVulDiscovered: ParseTimestamp(row, "last_vul_discovered", "lastVulDiscovered"));

    private static string PickFirstIp(JsonElement row)
    {
        var ip = ConnectSecureJsonReader.GetString(row, "ip");
        if (!string.IsNullOrWhiteSpace(ip))
            return ip;

        var combined = ConnectSecureJsonReader.GetString(row, "ip_addresses", "ipAddresses");
        if (string.IsNullOrWhiteSpace(combined))
            return "";

        var first = combined.Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return first.Length > 0 ? first[0] : "";
    }

    private static double GetDouble(JsonElement el, params string[] names)
    {
        foreach (var name in names)
        {
            if (!el.TryGetProperty(name, out var value))
                continue;
            if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var d))
                return d;
            if (value.ValueKind == JsonValueKind.String &&
                double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return d;
        }

        return 0d;
    }

    private static DateTimeOffset? ParseTimestamp(JsonElement el, params string[] names)
    {
        var text = ConnectSecureJsonReader.GetString(el, names);
        if (string.IsNullOrWhiteSpace(text))
            return null;
        return DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var dto)
            ? dto
            : null;
    }
}
