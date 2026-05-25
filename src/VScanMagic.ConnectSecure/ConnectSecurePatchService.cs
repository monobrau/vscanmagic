using System.Text.Json;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecurePatchService(
    ConnectSecureClient client,
    PatchActivityHistoryService patchActivityHistory,
    ConnectSecureCacheService cache)
{
    public async Task<IReadOnlyList<PatchableApplicationEntry>> GetPatchableApplicationsAsync(
        int companyId,
        bool patchableOnly = true,
        PatchApplicationLoadOptions? loadOptions = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        loadOptions ??= new PatchApplicationLoadOptions { PatchableOnly = patchableOnly };
        var effectivePatchableOnly = loadOptions.PatchableOnly && patchableOnly;

        var rows = await GetRemediationPlanRowsAsync(companyId, ct);

        return rows
            .Select(ParsePatchableApplication)
            .Where(entry => entry.SolutionId > 0 && !string.IsNullOrWhiteSpace(entry.Product))
            .Where(entry => !effectivePatchableOnly || entry.IsPatchable)
            .Where(entry => PatchCatalogHelper.MeetsSeverityFilter(entry.Severity, loadOptions.SeverityFilter))
            .Where(entry => !loadOptions.HideEndOfLife || !PatchCatalogHelper.IsEndOfLifeProductFix(entry.Fix))
            .OrderByDescending(entry => entry.AffectedAssets)
            .ThenBy(entry => entry.Product, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<IReadOnlyList<PatchAssetDetail>> GetPatchingAssetDetailsAsync(
        int companyId,
        int solutionId,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (solutionId <= 0)
            throw new ArgumentOutOfRangeException(nameof(solutionId), "Solution id is required.");

        var condition = $"solution_id={solutionId} and company_id={companyId}";
        if (cache.TryGetPatchHosts(companyId, solutionId, out var cached))
            return cached;

        var onlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: true, ct);
        var offlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: false, ct);
        await Task.WhenAll(onlineTask, offlineTask);

        var merged = PatchCatalogHelper.MergeAssetDetails((await onlineTask).Concat(await offlineTask)).ToList();
        cache.SetPatchHosts(companyId, solutionId, merged);
        return merged;
    }

    public Task<PatchOperationResult> PatchApplicationsNowAsync(
        ApplicationPatchRequest request,
        CancellationToken ct = default)
    {
        ValidatePatchRequest(request);
        var body = BuildPatchPayload(request, ConnectSecurePatchWhen.Now, scheduledAt: null);
        return InvokePatchAsync(request, "Application Patch", body, ct);
    }

    public Task<PatchOperationResult> PatchOsNowAsync(
        ApplicationPatchRequest request,
        CancellationToken ct = default)
    {
        var osRequest = request.Clone();
        osRequest.PatchType = ConnectSecurePatchType.Os;
        ValidatePatchRequest(osRequest);
        var body = BuildPatchPayload(osRequest, ConnectSecurePatchWhen.Now, scheduledAt: null);
        return InvokePatchAsync(osRequest, "OS Patch", body, ct);
    }

    public async Task<IReadOnlyList<PatchJobEntry>> GetCompanyJobsAsync(
        int companyId,
        bool patchJobsOnly = false,
        int limit = 50,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var localJobs = patchActivityHistory
            .GetEntries(companyId, limit)
            .Select(ToPatchJobEntry)
            .ToList();

        var rows = await FetchArrayAsync(
            "/r/company/jobs_view",
            new Dictionary<string, string>
            {
                ["condition"] = $"company_id={companyId}",
                ["limit"] = limit.ToString(),
                ["skip"] = "0",
                ["order_by"] = "updated desc"
            },
            ct);

        var remoteJobs = rows
            .Select(ParsePatchJob)
            .Where(job => !patchJobsOnly || IsPatchJobType(job.Type))
            .ToList();

        return localJobs
            .Concat(remoteJobs.Where(remote => localJobs.All(local =>
                !string.Equals(local.JobId, remote.JobId, StringComparison.OrdinalIgnoreCase))))
            .OrderByDescending(job => job.Updated ?? DateTimeOffset.MinValue)
            .Take(limit)
            .ToList();
    }

    public async Task<PatchVerificationResult> VerifyPatchActivityAsync(
        int companyId,
        string jobId,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (string.IsNullOrWhiteSpace(jobId))
            throw new ArgumentException("Job id is required.", nameof(jobId));

        var entry = patchActivityHistory.GetByJobId(companyId, jobId);
        if (entry is null)
            throw new InvalidOperationException("Patch activity entry not found.");

        var agentIds = entry.AgentIds?.Where(id => id > 0).Distinct().ToList() ?? [];
        if (agentIds.Count == 0)
        {
            var unverifiable = PatchCatalogHelper.BuildVerificationResult(jobId, [], []);
            UpdateVerificationEntry(entry, unverifiable);
            return unverifiable;
        }

        var targetFix = entry.TargetFix ?? "";
        var isEndOfLife = entry.IsEndOfLife;
        IReadOnlyList<PatchHostView> hostViews;

        if (entry.IsOsPatch)
        {
            var osPatch = new OsPendingPatchEntry(
                entry.Product ?? "",
                "",
                targetFix,
                entry.OsAssetIds?.Count ?? agentIds.Count,
                entry.OsAssetIds?.ToList() ?? []);
            var details = await GetOsPatchHostsAsync(companyId, osPatch, ct);
            hostViews = PatchCatalogHelper.BuildHostViews(details, targetFix, isEndOfLife);
        }
        else
        {
            var solutionIds = entry.SolutionIds?.Where(id => id > 0).Distinct().ToList() ?? [];
            if (solutionIds.Count == 0)
            {
                var unverifiable = PatchCatalogHelper.BuildVerificationResult(jobId, agentIds, []);
                UpdateVerificationEntry(entry, unverifiable);
                return unverifiable;
            }

            var details = await GetPatchingAssetDetailsForProductAsync(companyId, solutionIds, ct);
            hostViews = PatchCatalogHelper.BuildHostViews(details, targetFix, isEndOfLife);
        }

        var result = PatchCatalogHelper.BuildVerificationResult(jobId, agentIds, hostViews);
        UpdateVerificationEntry(entry, result);
        return result;
    }

    public async Task<IReadOnlyList<OsPendingPatchEntry>> GetOsPendingPatchesAsync(
        int companyId,
        int lookbackDays = 90,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var rows = await FetchAllPagesAsync(
            "/r/report_queries/os_pending_patches",
            new Dictionary<string, string>
            {
                ["num_days"] = lookbackDays.ToString(),
                ["condition"] = $"company_id={companyId}",
                ["order_by"] = "affected_assets desc"
            },
            ct);

        return rows
            .Where(row => MatchesCompany(row, companyId) || MatchesCompanyInIds(row, companyId))
            .Select(ParseOsPendingPatch)
            .Where(entry => entry.AffectedAssets > 0)
            .OrderByDescending(entry => entry.AffectedAssets)
            .ThenBy(entry => entry.OsName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<IReadOnlyList<PatchAssetDetail>> GetOsPatchHostsAsync(
        int companyId,
        OsPendingPatchEntry patch,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (patch.AssetIds.Count == 0)
            return [];

        var assetSet = patch.AssetIds.ToHashSet();
        var matched = await FetchOsHostsFromAgentsAsync(companyId, assetSet, ct);

        if (matched.Count < assetSet.Count)
        {
            var foundIds = matched
                .SelectMany(host => new[] { host.AssetId, host.AgentId })
                .Where(id => id > 0)
                .ToHashSet();
            var missing = assetSet.Where(id => !foundIds.Contains(id)).ToHashSet();
            if (missing.Count > 0)
            {
                var fromLightweight = await FetchLightweightAssetsMatchingAsync(companyId, missing, ct);
                matched.AddRange(fromLightweight);
            }
        }

        if (matched.Count == 0)
            return [];

        matched = await EnrichHostsWithAgentRegistryAsync(companyId, matched, ct);
        return PatchCatalogHelper.MergeAssetDetails(matched);
    }

    public async Task<IReadOnlyList<SuppressibleRemediationEntry>> GetSuppressibleRemediationsAsync(
        int companyId,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var rows = await GetRemediationPlanRowsAsync(companyId, ct);

        return rows
            .Select(row => new SuppressibleRemediationEntry(
                ConnectSecureJsonReader.GetInt(row, "solution_id", "solutionId") ?? 0,
                ConnectSecureJsonReader.GetString(row, "product"),
                ConnectSecureJsonReader.GetString(row, "fix"),
                ConnectSecureJsonReader.GetString(row, "severity"),
                ConnectSecureJsonReader.GetString(row, "remediation_action", "remediationAction"),
                ConnectSecureJsonReader.GetBool(row, "is_patchable", "isPatchable"),
                ConnectSecureJsonReader.GetInt(row, "affected_assets", "affectedAssets") ?? 0))
            .Where(entry => entry.SolutionId > 0 && !string.IsNullOrWhiteSpace(entry.Product))
            .OrderByDescending(entry => PatchCatalogHelper.SeverityRank(entry.Severity))
            .ThenByDescending(entry => entry.AffectedAssets)
            .ThenBy(entry => entry.Product, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<IReadOnlyList<SuppressibleProblemEntry>> GetSuppressibleProblemsAsync(
        int companyId,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var merged = new Dictionary<string, SuppressibleProblemEntry>(StringComparer.OrdinalIgnoreCase);

        var assetWiseTask = CollectSuppressibleProblemsAsync(
            merged,
            companyId,
            "/r/report_queries/asset_wise_vulnerabilities",
            AssetWiseQuery($"company_id={companyId}"),
            ct);

        var registryTask = CollectSuppressibleProblemsAsync(
            merged,
            companyId,
            "/r/report_queries/registry_problems_remediation",
            ReportQuery($"company_id={companyId} and is_suppressed=false and is_remediated = false"),
            ct);

        var networkTask = CollectSuppressibleProblemsAsync(
            merged,
            companyId,
            "/r/report_queries/application_vulnerabilities_net",
            ReportQuery($"company_id={companyId} and software_type='networksoftware' and unconfirmed = 'false'"),
            ct);

        await Task.WhenAll(assetWiseTask, registryTask, networkTask);

        return merged.Values
            .OrderByDescending(entry => entry.AffectedAssets)
            .ThenBy(entry => entry.ProblemName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<SuppressibleProblemEntry?> LookupProblemByNameAsync(
        int companyId,
        string problemName,
        string? source,
        IReadOnlyList<string>? hostIdentifiers = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var name = problemName.Trim();
        if (string.IsNullOrWhiteSpace(name))
            return null;

        if (CveReferenceHelper.IsCveOnlyProduct(name) || CveReferenceHelper.SplitCveIds(name).Count > 0)
        {
            var cveMatch = await LookupApplicationProblemFastAsync(companyId, name, hostIdentifiers, ct);
            if (cveMatch is not null)
                return cveMatch;
        }

        if (VulnerabilitySourceHelper.IsApplication(source))
            return await LookupApplicationProblemFastAsync(companyId, name, hostIdentifiers, ct);

        var escaped = EscapeConditionValue(name);
        if (VulnerabilitySourceHelper.IsRegistry(source))
        {
            return await LookupProblemFromEndpointAsync(
                companyId,
                name,
                "/r/report_queries/registry_problems_remediation",
                ReportQuery($"company_id={companyId} and is_suppressed=false and is_remediated = false and problem_name='{escaped}'"),
                ct);
        }

        if (VulnerabilitySourceHelper.IsNetwork(source))
        {
            return await LookupProblemFromEndpointAsync(
                companyId,
                name,
                "/r/report_queries/application_vulnerabilities_net",
                ReportQuery($"company_id={companyId} and software_type='networksoftware' and unconfirmed = 'false' and problem_name='{escaped}'"),
                ct);
        }

        var registryTask = LookupProblemFromEndpointAsync(
            companyId,
            name,
            "/r/report_queries/registry_problems_remediation",
            ReportQuery($"company_id={companyId} and is_suppressed=false and is_remediated = false and problem_name='{escaped}'"),
            ct);
        var networkTask = LookupProblemFromEndpointAsync(
            companyId,
            name,
            "/r/report_queries/application_vulnerabilities_net",
            ReportQuery($"company_id={companyId} and software_type='networksoftware' and unconfirmed = 'false' and problem_name='{escaped}'"),
            ct);
        var applicationTask = LookupApplicationProblemFastAsync(companyId, name, hostIdentifiers, ct);

        await Task.WhenAll(registryTask, networkTask, applicationTask);

        var registry = await registryTask;
        var network = await networkTask;
        var application = await applicationTask;

        return application ?? registry ?? network;
    }

    public async Task<SuppressibleProblemEntry?> LookupProblemByNameExhaustiveAsync(
        int companyId,
        string problemName,
        IReadOnlyList<string>? hostIdentifiers = null,
        CancellationToken ct = default)
    {
        var fast = await LookupProblemByNameAsync(companyId, problemName, source: null, hostIdentifiers, ct);
        if (fast is not null)
            return fast;

        var name = problemName.Trim();
        if (string.IsNullOrWhiteSpace(name))
            return null;

        return await LookupApplicationProblemDeepAsync(companyId, name, hostIdentifiers, ct);
    }

    public Task<PatchOperationResult> ScheduleApplicationPatchAsync(
        ScheduledApplicationPatchRequest request,
        CancellationToken ct = default)
    {
        ValidatePatchRequest(request);
        if (request.ScheduledAt <= DateTime.Now.AddMinutes(-1))
            throw new InvalidOperationException("Scheduled patch time must be in the future.");

        var body = BuildPatchPayload(request, ConnectSecurePatchWhen.Later, request.ScheduledAt);
        return InvokePatchAsync(request, "Scheduled Application Patch", body, ct);
    }

    internal static Dictionary<string, object?> BuildPatchPayload(
        ApplicationPatchRequest request,
        ConnectSecurePatchWhen patchWhen,
        DateTime? scheduledAt)
    {
        var body = new Dictionary<string, object?>
        {
            ["companies"] = new[] { request.CompanyId },
            ["patch_when"] = patchWhen == ConnectSecurePatchWhen.Now ? "now" : "later"
        };

        if (request.AssetIds.Count > 0)
            body["assets"] = request.AssetIds.ToArray();
        if (request.AgentIds.Count > 0)
            body["agents_id"] = request.AgentIds.ToArray();
        if (request.IncludedApplications.Count > 0)
            body["included_application"] = request.IncludedApplications.ToArray();
        if (request.ExcludedApplications.Count > 0)
            body["execluded_application"] = request.ExcludedApplications.ToArray();
        if (request.IncludeTags.Count > 0)
            body["include_tags"] = request.IncludeTags.ToArray();
        if (request.ExcludeTags.Count > 0)
            body["exclude_tags"] = request.ExcludeTags.ToArray();
        if (request.FromVersions.Count > 0)
            body["from_versions"] = new Dictionary<string, string>(request.FromVersions, StringComparer.Ordinal);

        if (request.PatchType == ConnectSecurePatchType.App)
            body["type"] = "application";

        if (request.TriggerReboot)
            body["trigger_reboot"] = true;

        if (patchWhen == ConnectSecurePatchWhen.Later)
        {
            if (scheduledAt is null)
                throw new InvalidOperationException("Scheduled patch time is required.");

            body["patch_type"] = request.PatchType == ConnectSecurePatchType.Os ? "os" : "app";
            body["date"] = new Dictionary<string, int>
            {
                ["days"] = scheduledAt.Value.Day,
                ["months"] = scheduledAt.Value.Month,
                ["years"] = scheduledAt.Value.Year
            };
            body["time"] = new Dictionary<string, int>
            {
                ["hours"] = scheduledAt.Value.Hour,
                ["minutes"] = scheduledAt.Value.Minute
            };
        }
        else if (request.PatchType == ConnectSecurePatchType.Os)
        {
            body["patch_type"] = "os";
        }

        return body;
    }

    private async Task<IReadOnlyList<PatchAssetDetail>> FetchRemediationPlanAssetDetailsAsync(
        string condition,
        bool onlineOnly,
        CancellationToken ct)
    {
        var rows = await FetchAllPagesAsync(
            "/r/report_queries/remediation_plan_asset_details",
            new Dictionary<string, string>
            {
                ["having"] = onlineOnly ? "true" : "false",
                ["condition"] = condition
            },
            ct);

        return rows
            .Select(ParseRemediationPlanAssetDetail)
            .Where(detail => detail.AgentId > 0 || detail.AssetId > 0 || !string.IsNullOrWhiteSpace(detail.HostName))
            .ToList();
    }

    public async Task<IReadOnlyList<PatchAssetDetail>> GetPatchingAssetDetailsForProductAsync(
        int companyId,
        IReadOnlyList<int> solutionIds,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (solutionIds.Count == 0)
            return [];

        var distinctSolutionIds = solutionIds.Distinct().ToList();
        using var gate = new SemaphoreSlim(3);
        var detailTasks = distinctSolutionIds.Select(async solutionId =>
        {
            await gate.WaitAsync(ct);
            try
            {
                return await GetPatchingAssetDetailsAsync(companyId, solutionId, ct);
            }
            finally
            {
                gate.Release();
            }
        }).ToList();

        var merged = (await Task.WhenAll(detailTasks)).SelectMany(details => details).ToList();

        var combined = PatchCatalogHelper.MergeAssetDetails(merged);
        return await EnrichHostsWithAgentRegistryAsync(companyId, combined.ToList(), ct);
    }

    private async Task<PatchOperationResult> InvokePatchAsync(
        ApplicationPatchRequest request,
        string jobType,
        Dictionary<string, object?> body,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/patch_now",
            body: body,
            ct: ct);

        var success = ConnectSecureJsonReader.GetBool(response, "status");
        var message = ConnectSecureJsonReader.GetString(response, "message");
        if (string.IsNullOrWhiteSpace(message))
            message = success ? "Patch request accepted." : "ConnectSecure rejected the patch request.";

        if (!success)
            throw new InvalidOperationException(message);

        cache.InvalidateCompany(request.CompanyId);

        var jobId = RecordPatchActivity(request, jobType, message);
        return new PatchOperationResult(true, message, jobId);
    }

    private string RecordPatchActivity(ApplicationPatchRequest request, string jobType, string message)
    {
        var jobId = Guid.NewGuid().ToString("N");
        var hostNames = request.TargetHostNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var hostSummary = hostNames.Count switch
        {
            0 => $"{request.AgentIds.Count} host(s)",
            1 => hostNames[0],
            <= 3 => string.Join(", ", hostNames),
            _ => $"{string.Join(", ", hostNames.Take(2))} +{hostNames.Count - 2} more"
        };

        var products = request.IncludedApplications
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var productSummary = products.Count switch
        {
            0 => "Patch request",
            1 => products[0],
            _ => string.Join(", ", products)
        };

        var description = request.PatchType == ConnectSecurePatchType.Os && request.TriggerReboot
            ? $"{productSummary} on {hostSummary} (reboot requested)"
            : $"{productSummary} on {hostSummary}";

        var product = products.Count == 1 ? products[0] : products.FirstOrDefault() ?? "";

        patchActivityHistory.Record(new PatchActivityEntry(
            request.CompanyId,
            jobId,
            jobType,
            "Submitted",
            description,
            hostNames.Count == 1 ? hostNames[0] : hostSummary,
            null,
            DateTimeOffset.Now,
            message,
            request.AgentIds.Where(id => id > 0).Distinct().ToList(),
            request.SolutionIds.Where(id => id > 0).Distinct().ToList(),
            product,
            request.TargetFix,
            request.PatchType == ConnectSecurePatchType.Os,
            request.IsEndOfLife,
            request.OsAssetIds.Where(id => id > 0).Distinct().ToList()));

        return jobId;
    }

    private void UpdateVerificationEntry(PatchActivityEntry entry, PatchVerificationResult result)
    {
        patchActivityHistory.UpdateEntry(entry with
        {
            Status = result.Status,
            VerificationSummary = result.Summary,
            VerifiedAt = result.VerifiedAt
        });
    }

    private static PatchJobEntry ToPatchJobEntry(PatchActivityEntry entry)
    {
        var canVerify = entry.AgentIds is { Count: > 0 } &&
                        (entry.SolutionIds is { Count: > 0 } || entry.IsOsPatch);
        var description = string.IsNullOrWhiteSpace(entry.ConnectSecureMessage)
            ? entry.Description
            : $"{entry.Description} — {entry.ConnectSecureMessage}";
        if (!string.IsNullOrWhiteSpace(entry.VerificationSummary))
            description = $"{description} — {entry.VerificationSummary}";

        return new PatchJobEntry(
            entry.JobId,
            entry.Type,
            entry.Status,
            description,
            entry.HostName,
            entry.AgentIp,
            entry.VerifiedAt ?? entry.RequestedAt,
            IsLocal: true,
            CanVerify: canVerify,
            VerificationSummary: entry.VerificationSummary);
    }

    private static void ValidatePatchRequest(ApplicationPatchRequest request)
    {
        if (request.CompanyId <= 0)
            throw new InvalidOperationException("Company id is required.");
        if (request.PatchType != ConnectSecurePatchType.Os && request.IncludedApplications.Count == 0)
            throw new InvalidOperationException("At least one application must be included in the patch.");
        if (request.AssetIds.Count == 0 && request.AgentIds.Count == 0)
            throw new InvalidOperationException("Select at least one asset or agent to patch.");
    }

    private static bool IsPatchJobType(string? type)
    {
        if (string.IsNullOrWhiteSpace(type))
            return false;

        return type.Contains("patch", StringComparison.OrdinalIgnoreCase) ||
               type.Contains("remediation", StringComparison.OrdinalIgnoreCase);
    }

    private static PatchJobEntry ParsePatchJob(JsonElement row)
    {
        var updatedText = ConnectSecureJsonReader.GetString(row, "updated", "created");
        DateTimeOffset? updated = null;
        if (!string.IsNullOrWhiteSpace(updatedText) && DateTimeOffset.TryParse(updatedText, out var parsed))
            updated = parsed.ToLocalTime();

        return new PatchJobEntry(
            ConnectSecureJsonReader.GetString(row, "job_id", "jobId", "id"),
            ConnectSecureJsonReader.GetString(row, "type"),
            ConnectSecureJsonReader.GetString(row, "status"),
            ConnectSecureJsonReader.GetString(row, "description", "name"),
            ConnectSecureJsonReader.GetString(row, "agent_host_name", "agentHostName", "host_name", "hostName"),
            ConnectSecureJsonReader.GetString(row, "agent_ip", "agentIp", "ip"),
            updated);
    }

    private static OsPendingPatchEntry ParseOsPendingPatch(JsonElement row)
    {
        var assetIds = ExtractIntArray(row, "asset_ids", "assetIds");
        var affected = ConnectSecureJsonReader.GetInt(row, "affected_assets", "affectedAssets") ?? assetIds.Count;
        if (assetIds.Count > 0)
            affected = Math.Min(affected, assetIds.Count);

        return new OsPendingPatchEntry(
            ConnectSecureJsonReader.GetString(row, "os_name", "osName"),
            ConnectSecureJsonReader.GetString(row, "os_version", "osVersion"),
            ConnectSecureJsonReader.GetString(row, "fix"),
            affected,
            assetIds);
    }

    private static bool LightweightRowMatchesAssetIds(JsonElement row, HashSet<int> assetSet)
    {
        foreach (var name in new[] { "asset_id", "assetId", "id" })
        {
            if (!row.TryGetProperty(name, out var value))
                continue;

            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var n) && assetSet.Contains(n))
                return true;

            if (value.ValueKind == JsonValueKind.String &&
                int.TryParse(value.GetString(), out n) &&
                assetSet.Contains(n))
                return true;
        }

        return false;
    }

    private async Task<List<JsonElement>> GetRemediationPlanRowsAsync(int companyId, CancellationToken ct)
    {
        if (cache.TryGetRemediationPlan(companyId, out var cached))
            return cached;

        var rows = await FetchRemediationPlanForCompanyAsync(
            companyId,
            ReportQuery($"company_id={companyId}"),
            ct);
        cache.SetRemediationPlan(companyId, rows);
        return rows;
    }

    public void InvalidateCompanyCache(int companyId) => cache.InvalidateCompany(companyId);

    private async Task<List<PatchAssetDetail>> EnrichHostsWithAgentRegistryAsync(
        int companyId,
        List<PatchAssetDetail> hosts,
        CancellationToken ct)
    {
        if (hosts.Count == 0)
            return hosts;

        var neededAgentIds = hosts
            .Where(host => host.AgentId > 0)
            .Select(host => host.AgentId)
            .ToHashSet();
        var neededAssetIds = hosts
            .Where(host => host.AgentId <= 0 && host.AssetId > 0)
            .Select(host => host.AssetId)
            .ToHashSet();
        var neededHostNames = hosts
            .Where(host => host.AgentId <= 0 && !string.IsNullOrWhiteSpace(host.HostName))
            .Select(host => host.HostName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var byAgentId = new Dictionary<int, PatchAssetDetail>();
        var byAssetId = new Dictionary<int, PatchAssetDetail>();
        var byHostName = new Dictionary<string, PatchAssetDetail>(StringComparer.OrdinalIgnoreCase);

        if (neededAgentIds.Count > 0)
        {
            using var gate = new SemaphoreSlim(5);
            var agentTasks = neededAgentIds.Select(async agentId =>
            {
                if (cache.TryGetAgentDetail(agentId, out var cachedAgent))
                    return (agentId, cachedAgent);

                await gate.WaitAsync(ct);
                try
                {
                    if (cache.TryGetAgentDetail(agentId, out cachedAgent))
                        return (agentId, cachedAgent);

                    var agent = await FetchAgentRegistryDetailAsync(agentId, ct);
                    if (agent is not null)
                        cache.SetAgentDetail(agentId, agent);
                    return (agentId, agent);
                }
                finally
                {
                    gate.Release();
                }
            }).ToList();

            foreach (var (agentId, agent) in await Task.WhenAll(agentTasks))
            {
                if (agent is null)
                    continue;

                byAgentId.TryAdd(agentId, agent);
                neededAgentIds.Remove(agentId);
            }
        }

        for (var page = 0;
             page < ConnectSecurePagedQuery.MaxPages &&
             (neededAgentIds.Count > 0 || neededAssetIds.Count > 0 || neededHostNames.Count > 0);
             page++)
        {
            ct.ThrowIfCancellationRequested();

            var agents = await FetchArrayAsync(
                "/r/company/agents",
                ConnectSecureCompanyReviewService.CompanyQuery(
                    companyId,
                    limit: ConnectSecurePagedQuery.PageSize,
                    skip: page * ConnectSecurePagedQuery.PageSize),
                ct);

            if (agents.Count == 0)
                break;

            foreach (var row in agents)
            {
                var agent = ParseAgentAssetDetail(row);
                if (agent.AgentId > 0)
                    cache.SetAgentDetail(agent.AgentId, agent);

                if (agent.AgentId > 0 && neededAgentIds.Contains(agent.AgentId))
                {
                    byAgentId.TryAdd(agent.AgentId, agent);
                    neededAgentIds.Remove(agent.AgentId);
                }

                if (agent.AssetId > 0 && neededAssetIds.Contains(agent.AssetId))
                {
                    byAssetId.TryAdd(agent.AssetId, agent);
                    neededAssetIds.Remove(agent.AssetId);
                }

                if (!string.IsNullOrWhiteSpace(agent.HostName) && neededHostNames.Contains(agent.HostName))
                {
                    byHostName.TryAdd(agent.HostName, agent);
                    neededHostNames.Remove(agent.HostName);
                }
            }

            if (agents.Count < ConnectSecurePagedQuery.PageSize &&
                neededAgentIds.Count == 0 &&
                neededAssetIds.Count == 0 &&
                neededHostNames.Count == 0)
                break;
        }

        var enriched = new List<PatchAssetDetail>();
        foreach (var host in hosts)
        {
            if (host.AgentId > 0 && byAgentId.TryGetValue(host.AgentId, out var byAgent))
            {
                enriched.Add(MergeHostWithAgent(host, byAgent));
                continue;
            }

            if (host.AssetId > 0 && byAssetId.TryGetValue(host.AssetId, out var byAsset))
            {
                enriched.Add(MergeHostWithAgent(host, byAsset));
                continue;
            }

            if (!string.IsNullOrWhiteSpace(host.HostName) &&
                byHostName.TryGetValue(host.HostName, out var byName))
            {
                enriched.Add(MergeHostWithAgent(host, byName));
                continue;
            }

            enriched.Add(host);
        }

        return enriched;
    }

    private async Task<PatchAssetDetail?> FetchAgentRegistryDetailAsync(int agentId, CancellationToken ct)
    {
        if (agentId <= 0)
            return null;

        try
        {
            var response = await client.InvokeAuthenticatedAsync(
                HttpMethod.Get,
                $"/r/company/agents/{agentId}",
                ct: ct);
            if (!response.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Object)
                return null;

            return ParseAgentAssetDetail(data);
        }
        catch
        {
            return null;
        }
    }

    private static PatchAssetDetail MergeHostWithAgent(PatchAssetDetail host, PatchAssetDetail agent) =>
        host with
        {
            AgentId = agent.AgentId > 0 ? agent.AgentId : host.AgentId,
            AssetId = host.AssetId > 0 ? host.AssetId : agent.AssetId,
            OnlineStatus = agent.OnlineStatus,
            RegisteredHostName = string.IsNullOrWhiteSpace(agent.HostName) ? host.RegisteredHostName : agent.HostName,
            AgentType = agent.AgentType ?? host.AgentType,
            AgentVersion = agent.AgentVersion ?? host.AgentVersion,
            LastPingTime = agent.LastPingTime ?? host.LastPingTime,
            Ip = string.IsNullOrWhiteSpace(host.Ip) ? agent.Ip : host.Ip
        };

    private async Task<List<PatchAssetDetail>> FetchOsHostsFromAgentsAsync(
        int companyId,
        HashSet<int> assetSet,
        CancellationToken ct)
    {
        var matched = new List<PatchAssetDetail>();
        var remaining = new HashSet<int>(assetSet);

        for (var page = 0;
             page < ConnectSecurePagedQuery.MaxPages && remaining.Count > 0;
             page++)
        {
            ct.ThrowIfCancellationRequested();

            var agents = await FetchArrayAsync(
                "/r/company/agents",
                ConnectSecureCompanyReviewService.CompanyQuery(
                    companyId,
                    limit: ConnectSecurePagedQuery.PageSize,
                    skip: page * ConnectSecurePagedQuery.PageSize),
                ct);

            if (agents.Count == 0)
                break;

            foreach (var agent in agents.Where(row => LightweightRowMatchesAssetIds(row, remaining)))
            {
                var detail = ParseAgentAssetDetail(agent);
                matched.Add(detail);
                if (detail.AssetId > 0)
                    remaining.Remove(detail.AssetId);
                if (detail.AgentId > 0)
                    remaining.Remove(detail.AgentId);
            }

            if (agents.Count < ConnectSecurePagedQuery.PageSize)
                break;
        }

        return matched;
    }

    private async Task<List<PatchAssetDetail>> FetchLightweightAssetsMatchingAsync(
        int companyId,
        HashSet<int> assetSet,
        CancellationToken ct)
    {
        var matched = new List<PatchAssetDetail>();
        var remaining = new HashSet<int>(assetSet);

        for (var page = 0;
             page < ConnectSecurePagedQuery.MaxPages && remaining.Count > 0;
             page++)
        {
            ct.ThrowIfCancellationRequested();

            var rows = await FetchArrayAsync(
                "/r/report_queries/lightweight_assets",
                new Dictionary<string, string>
                {
                    ["company_id"] = companyId.ToString(),
                    ["limit"] = ConnectSecurePagedQuery.PageSize.ToString(),
                    ["skip"] = (page * ConnectSecurePagedQuery.PageSize).ToString()
                },
                ct);

            if (rows.Count == 0)
                break;

            foreach (var row in rows.Where(row => LightweightRowMatchesAssetIds(row, remaining)))
            {
                var detail = ParseLightweightAssetDetail(row);
                matched.Add(detail);
                if (detail.AssetId > 0)
                    remaining.Remove(detail.AssetId);
                if (detail.AgentId > 0)
                    remaining.Remove(detail.AgentId);
            }

            if (rows.Count < ConnectSecurePagedQuery.PageSize)
                break;
        }

        return matched;
    }

    private static PatchAssetDetail ParseAgentAssetDetail(JsonElement row)
    {
        var agentId = ConnectSecureJsonReader.GetInt(row, "id", "agent_id", "agentId") ?? 0;
        var assetId = ConnectSecureJsonReader.GetInt(row, "asset_id", "assetId") ?? 0;
        var lastPing = ConnectSecureJsonReader.GetString(row, "last_ping_time", "lastPingTime");
        var lastReported = ConnectSecureJsonReader.GetString(row, "last_reported", "lastReported", "last_scanned_time");
        var online = AgentConnectivityHelper.IsOnlineFromAgentTimestamps(lastPing, lastReported);

        return new PatchAssetDetail(
            assetId,
            ConnectSecureJsonReader.GetString(row, "ip"),
            ConnectSecureJsonReader.GetString(row, "host_name", "hostName", "name"),
            agentId,
            online,
            [ConnectSecureJsonReader.GetString(row, "os_name", "osName")],
            [ConnectSecureJsonReader.GetString(row, "os_version", "osVersion")],
            [],
            AgentType: ConnectSecureJsonReader.GetString(row, "agent_type", "agentType"),
            AgentVersion: ConnectSecureJsonReader.GetString(row, "agent_version", "agentVersion", "version"),
            LastPingTime: lastPing);
    }

    private static bool MatchesCompanyInIds(JsonElement row, int companyId)
    {
        if (row.TryGetProperty("company_ids", out var ids) || row.TryGetProperty("companyIds", out ids))
        {
            if (ids.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in ids.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.Number && item.TryGetInt32(out var n) && n == companyId)
                        return true;
                    if (item.ValueKind == JsonValueKind.String &&
                        int.TryParse(item.GetString(), out n) &&
                        n == companyId)
                        return true;
                }

                return false;
            }
        }

        return MatchesCompany(row, companyId);
    }

    private static PatchAssetDetail ParseLightweightAssetDetail(JsonElement row) =>
        new(
            ConnectSecureJsonReader.GetInt(row, "asset_id", "assetId", "id") ?? 0,
            ConnectSecureJsonReader.GetString(row, "ip"),
            ConnectSecureJsonReader.GetString(row, "host_name", "hostName"),
            ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
            ConnectSecureJsonReader.GetBool(row, "online_status", "onlineStatus"),
            [ConnectSecureJsonReader.GetString(row, "os_name", "osName", "operating_system")],
            [ConnectSecureJsonReader.GetString(row, "os_version", "osVersion")],
            []);

    private static bool MatchesCompany(JsonElement row, int companyId) =>
        (ConnectSecureJsonReader.GetInt(row, "company_id", "companyId") ?? 0) == companyId;

    private static PatchableApplicationEntry ParsePatchableApplication(JsonElement row) =>
        new(
            ConnectSecureJsonReader.GetInt(row, "solution_id", "solutionId") ?? 0,
            ConnectSecureJsonReader.GetString(row, "product"),
            ConnectSecureJsonReader.GetString(row, "fix"),
            ConnectSecureJsonReader.GetBool(row, "is_patchable", "isPatchable"),
            ConnectSecureJsonReader.GetInt(row, "affected_assets", "affectedAssets") ?? 0,
            ExtractIntArray(row, "asset_ids", "assetIds"),
            ConnectSecureJsonReader.GetString(row, "severity"),
            ConnectSecureJsonReader.GetString(row, "remediation_action", "remediationAction"));

    private static SuppressibleProblemEntry ParseSuppressibleProblem(JsonElement row)
    {
        var problemName = ConnectSecureJsonReader.GetString(row, "problem_name", "problemName", "name");
        if (string.IsNullOrWhiteSpace(problemName))
            problemName = ConnectSecureJsonReader.GetString(row, "software_name", "softwareName");

        return new SuppressibleProblemEntry(
            ConnectSecureJsonReader.GetInt(row, "problem_id", "problemId")
                ?? ConnectSecureJsonReader.GetInt(row, "id")
                ?? 0,
            problemName,
            ConnectSecureJsonReader.GetInt(row, "affected_assets", "affectedAssets") ?? 0);
    }

    private static PatchAssetDetail ParseRemediationPlanAssetDetail(JsonElement row)
    {
        var productName = ConnectSecureJsonReader.GetString(row, "name", "product");
        var applicationNames = string.IsNullOrWhiteSpace(productName) ? [] : new[] { productName };

        return new PatchAssetDetail(
            ConnectSecureJsonReader.GetInt(row, "id", "asset_id", "assetId") ?? 0,
            ConnectSecureJsonReader.GetString(row, "ip"),
            ConnectSecureJsonReader.GetString(row, "host_name", "hostName"),
            ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
            ConnectSecureJsonReader.GetBool(row, "online_status", "onlineStatus"),
            applicationNames,
            PatchCatalogHelper.NormalizeVersions(ExtractVersionInstallDates(row)),
            ExtractStringArray(row, "uninstall_string", "uninstallString"));
    }

    private static IReadOnlyList<string> ExtractVersionInstallDates(JsonElement row)
    {
        if (!row.TryGetProperty("version_install_date", out var value) &&
            !row.TryGetProperty("versionInstallDate", out value))
            return ExtractStringArray(row, "version");

        if (value.ValueKind != JsonValueKind.Array)
            return [];

        var versions = new List<string>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Object)
            {
                var version = ConnectSecureJsonReader.GetString(item, "version");
                if (!string.IsNullOrWhiteSpace(version))
                    versions.Add(version);
            }
            else if (item.ValueKind == JsonValueKind.String)
            {
                var text = item.GetString();
                if (!string.IsNullOrWhiteSpace(text))
                    versions.Add(text);
            }
        }

        return versions;
    }

    private static IReadOnlyList<int> ExtractIntArray(JsonElement row, params string[] names)
    {
        foreach (var name in names)
        {
            if (!row.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Array)
                continue;

            var ids = new List<int>();
            foreach (var item in value.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Number && item.TryGetInt32(out var n))
                    ids.Add(n);
                else if (item.ValueKind == JsonValueKind.String && int.TryParse(item.GetString(), out n))
                    ids.Add(n);
            }

            return ids;
        }

        return [];
    }

    private static IReadOnlyList<string> ExtractStringArray(JsonElement row, params string[] names)
    {
        foreach (var name in names)
        {
            if (!row.TryGetProperty(name, out var value))
                continue;

            if (value.ValueKind == JsonValueKind.String)
            {
                var text = value.GetString();
                return string.IsNullOrWhiteSpace(text) ? [] : [text];
            }

            if (value.ValueKind != JsonValueKind.Array)
                continue;

            return value.EnumerateArray()
                .Select(item => item.ValueKind == JsonValueKind.String ? item.GetString() ?? "" : item.ToString())
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
        }

        return [];
    }

    private async Task<List<JsonElement>> FetchArrayAsync(
        string endpoint,
        IReadOnlyDictionary<string, string> query,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, endpoint, query, ct: ct);
        return ConnectSecureJsonReader.ExtractDataArray(response);
    }

    private Task<List<JsonElement>> FetchRemediationPlanForCompanyAsync(
        int companyId,
        Dictionary<string, string> baseQuery,
        CancellationToken ct) =>
        ConnectSecurePagedQuery.FetchCompanyScopedPagesAsync(
            (query, token) => FetchArrayAsync("/r/report_queries/remediation_plan_by_company", query, token),
            baseQuery,
            companyId,
            ct);

    private Task<List<JsonElement>> FetchAllPagesAsync(
        string endpoint,
        Dictionary<string, string> baseQuery,
        CancellationToken ct) =>
        ConnectSecurePagedQuery.FetchAllPagesAsync(
            (query, token) => FetchArrayAsync(endpoint, query, token),
            baseQuery,
            ct);

    private static Dictionary<string, string> BuildApplicationVulnQuery(int companyId) =>
        BuildVulnerabilitiesDetailsQuery(companyId);

    private static Dictionary<string, string> BuildVulnerabilitiesDetailsQuery(
        int? companyId = null,
        string? escapedProblemName = null,
        int limit = 25)
    {
        var parts = new List<string>();
        if (companyId is > 0)
            parts.Add($"company_ids:{companyId.Value}");
        if (!string.IsNullOrWhiteSpace(escapedProblemName))
            parts.Add($"problem_name='{escapedProblemName}'");

        var query = new Dictionary<string, string>
        {
            ["order_by"] = "affected_assets desc",
            ["limit"] = limit.ToString(),
            ["skip"] = "0"
        };

        if (parts.Count > 0)
            query["condition"] = string.Join(" and ", parts);

        return query;
    }

    private static string EscapeConditionValue(string value) =>
        value.Replace("'", "\\'", StringComparison.Ordinal);

    private async Task CollectSuppressibleProblemsAsync(
        IDictionary<string, SuppressibleProblemEntry> merged,
        int companyId,
        string endpoint,
        Dictionary<string, string> query,
        CancellationToken ct)
    {
        var rows = await FetchAllPagesAsync(endpoint, query, ct);
        foreach (var row in rows)
        {
            if (!MatchesCompanyInIds(row, companyId))
                continue;

            ConnectSecureSuppressLookup.MergeProblem(merged, ParseSuppressibleProblem(row));
        }
    }

    private async Task<SuppressibleProblemEntry?> LookupApplicationProblemFastAsync(
        int companyId,
        string problemName,
        IReadOnlyList<string>? hostIdentifiers,
        CancellationToken ct)
    {
        var escaped = EscapeConditionValue(problemName);

        // asset_wise includes problem_id; vulnerabilities_details rows often do not.
        var direct = await TryFindProblemInAssetWiseAsync(
            companyId,
            problemName,
            AssetWiseQuery($"company_id={companyId} and problem_name='{escaped}'").WithLookupLimit(25),
            ct);
        if (direct is not null)
            return direct;

        if (hostIdentifiers is { Count: > 0 })
        {
            var hostMatch = await LookupApplicationProblemByHostsAsync(companyId, problemName, hostIdentifiers, ct);
            if (hostMatch is not null)
                return hostMatch;
        }

        return null;
    }

    private async Task<SuppressibleProblemEntry?> LookupApplicationProblemByHostsAsync(
        int companyId,
        string problemName,
        IReadOnlyList<string> hostIdentifiers,
        CancellationToken ct)
    {
        var escaped = EscapeConditionValue(problemName);
        foreach (var host in hostIdentifiers
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Select(h => h.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(5))
        {
            ct.ThrowIfCancellationRequested();
            var escapedHost = EscapeConditionValue(host);
            var hostField = host.Contains('.', StringComparison.Ordinal) ? "ip" : "host_name";
            var match = await TryFindProblemInAssetWiseAsync(
                companyId,
                problemName,
                AssetWiseQuery($"company_id={companyId} and {hostField}='{escapedHost}' and problem_name='{escaped}'")
                    .WithLookupLimit(25),
                ct);
            if (match is not null)
                return match;
        }

        return null;
    }

    private async Task<SuppressibleProblemEntry?> TryFindProblemInAssetWiseAsync(
        int companyId,
        string problemName,
        Dictionary<string, string> query,
        CancellationToken ct)
    {
        var rows = await FetchArrayAsync("/r/report_queries/asset_wise_vulnerabilities", query, ct);
        return ConnectSecureSuppressLookup.FindProblemMatch(
            rows, companyId, problemName, MatchesCompanyInIds, ParseSuppressibleProblem);
    }

    private async Task<SuppressibleProblemEntry?> LookupApplicationProblemDeepAsync(
        int companyId,
        string problemName,
        IReadOnlyList<string>? hostIdentifiers,
        CancellationToken ct)
    {
        var fast = await LookupApplicationProblemFastAsync(companyId, problemName, hostIdentifiers, ct);
        if (fast is not null)
            return fast;

        for (var page = 0; page < 3; page++)
        {
            var query = AssetWiseQuery($"company_id={companyId}");
            query["limit"] = ConnectSecurePagedQuery.PageSize.ToString();
            query["skip"] = (page * ConnectSecurePagedQuery.PageSize).ToString();

            var rows = await FetchArrayAsync("/r/report_queries/asset_wise_vulnerabilities", query, ct);
            if (rows.Count == 0)
                break;

            var match = ConnectSecureSuppressLookup.FindProblemMatch(
                rows, companyId, problemName, MatchesCompanyInIds, ParseSuppressibleProblem);
            if (match is not null)
                return match;

            if (rows.Count < ConnectSecurePagedQuery.PageSize)
                break;
        }

        return null;
    }

    private async Task<SuppressibleProblemEntry?> LookupProblemFromEndpointAsync(
        int companyId,
        string problemName,
        string endpoint,
        Dictionary<string, string> query,
        CancellationToken ct)
    {
        var rows = await FetchArrayAsync(endpoint, query.WithLookupLimit(), ct);
        return ConnectSecureSuppressLookup.FindProblemMatch(
            rows, companyId, problemName, MatchesCompanyInIds, ParseSuppressibleProblem);
    }

    private static Dictionary<string, string> ReportQuery(string condition) =>
        new()
        {
            ["condition"] = condition,
            ["order_by"] = "affected_assets desc"
        };

    private static Dictionary<string, string> AssetWiseQuery(string condition) =>
        new()
        {
            ["condition"] = condition,
            ["order_by"] = "severity desc"
        };
}
