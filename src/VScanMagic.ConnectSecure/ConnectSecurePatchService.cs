using System.Text.Json;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecurePatchService(
    ConnectSecureClient client,
    PatchActivityHistoryService patchActivityHistory,
    ConnectSecureCacheService cache,
    ConnectSecureRemediationService remediationService,
    ConnectSecureScanService scanService)
{
    public async Task<IReadOnlyList<PatchableApplicationEntry>> GetPatchableApplicationsAsync(
        int companyId,
        bool patchableOnly = true,
        PatchApplicationLoadOptions? loadOptions = null,
        bool forceRefresh = false,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        loadOptions ??= new PatchApplicationLoadOptions { PatchableOnly = patchableOnly };
        var effectivePatchableOnly = loadOptions.PatchableOnly && patchableOnly;

        if (forceRefresh)
        {
            remediationService.Invalidate(companyId);
            cache.InvalidatePatchHosts(companyId);
        }

        var dataset = await remediationService.GetRemediationDatasetAsync(companyId, forceRefresh: forceRefresh, ct: ct);

        return dataset.Records
            .Where(r => r.IsSoftwarePatch)
            .Where(r => r.SolutionId > 0 && !string.IsNullOrWhiteSpace(r.Product))
            .GroupBy(r => r.SolutionId)
            .Select(BuildPatchableApplicationEntry)
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

        if (cache.TryGetPatchHosts(companyId, solutionId, out var cached))
            return cached;

        var dataset = await remediationService.GetRemediationDatasetAsync(companyId, ct: ct);
        var solutionRecords = dataset.Records.Where(r => r.SolutionId == solutionId).ToList();
        var productName = solutionRecords.FirstOrDefault()?.Product ?? "";
        var rawHosts = DeduplicateRemediationHosts(solutionRecords);

        var merged = PatchCatalogHelper.MergeAssetDetails(rawHosts).ToList();
        if (!string.IsNullOrWhiteSpace(productName))
        {
            var assetDetails = await GetOrFetchProductRemediationAssetDetailsAsync(companyId, productName, ct);
            merged = MergeInstalledVersions(merged, assetDetails, productName);
        }

        var enriched = await EnrichHostsWithAgentRegistryAsync(companyId, merged, ct);
        cache.SetPatchHosts(companyId, solutionId, enriched);
        return enriched;
    }

    private async Task<List<PatchAssetDetail>> GetOrFetchProductRemediationAssetDetailsAsync(
        int companyId,
        string productName,
        CancellationToken ct)
    {
        var normalizedProduct = productName.Trim();
        if (string.IsNullOrWhiteSpace(normalizedProduct))
            return [];

        if (cache.TryGetProductRemediationAssetDetails(companyId, normalizedProduct, out var cached))
            return cached;

        var escapedProduct = EscapeConditionValue(normalizedProduct);
        var condition = $"company_id={companyId} and name='{escapedProduct}'";
        var onlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: true, ct);
        var offlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: false, ct);
        await Task.WhenAll(onlineTask, offlineTask);

        var combined = (await onlineTask).Concat(await offlineTask).ToList();
        var deduped = DedupeAssetDetailsByHostAndProduct(combined.Where(IsUsableAssetDetailRow));

        if (deduped.Count == 0)
        {
            var companyWide = await GetOrFetchCompanyRemediationAssetDetailsAsync(companyId, ct);
            deduped = companyWide
                .Where(detail => detail.ApplicationNames.Any(name =>
                    PatchCatalogHelper.ProductNamesMatch(name, normalizedProduct)))
                .ToList();
        }

        cache.SetProductRemediationAssetDetails(companyId, normalizedProduct, deduped);
        return deduped;
    }

    private async Task<List<PatchAssetDetail>> GetOrFetchCompanyRemediationAssetDetailsAsync(
        int companyId,
        CancellationToken ct)
    {
        if (cache.TryGetRemediationAssetDetails(companyId, out var cached))
            return cached;

        var condition = $"company_id={companyId}";
        var onlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: true, ct);
        var offlineTask = FetchRemediationPlanAssetDetailsAsync(condition, onlineOnly: false, ct);
        await Task.WhenAll(onlineTask, offlineTask);

        var combined = (await onlineTask).Concat(await offlineTask).ToList();
        var deduped = DedupeAssetDetailsByHostAndProduct(combined.Where(IsUsableAssetDetailRow));
        cache.SetRemediationAssetDetails(companyId, deduped);
        return deduped;
    }

    private static bool IsUsableAssetDetailRow(PatchAssetDetail detail) =>
        detail.AgentId > 0 &&
        (!string.IsNullOrWhiteSpace(detail.HostName) || detail.Versions.Count > 0);

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

    private static List<PatchAssetDetail> DedupeAssetDetailsByHostAndProduct(IEnumerable<PatchAssetDetail> details)
    {
        var deduped = new Dictionary<string, PatchAssetDetail>(StringComparer.OrdinalIgnoreCase);
        foreach (var detail in details)
        {
            var product = detail.ApplicationNames.FirstOrDefault(name => !string.IsNullOrWhiteSpace(name)) ?? "";
            var key = $"{detail.AgentId}:{detail.AssetId}:{product}";
            if (string.IsNullOrEmpty(key.Trim(':')))
                continue;

            if (!deduped.TryGetValue(key, out var existing) ||
                (!existing.OnlineStatus && detail.OnlineStatus))
            {
                deduped[key] = detail;
            }
        }

        return deduped.Values.ToList();
    }

    private static List<PatchAssetDetail> MergeInstalledVersions(
        IReadOnlyList<PatchAssetDetail> hosts,
        IReadOnlyList<PatchAssetDetail> assetDetails,
        string productName)
    {
        if (hosts.Count == 0 || assetDetails.Count == 0)
            return hosts.ToList();

        var productDetails = assetDetails
            .Where(detail => detail.ApplicationNames.Any(name =>
                PatchCatalogHelper.ProductNamesMatch(name, productName)))
            .ToList();

        var byAgentId = productDetails
            .Where(detail => detail.AgentId > 0)
            .GroupBy(detail => detail.AgentId)
            .ToDictionary(group => group.Key, group => group.First());
        var byAssetId = productDetails
            .Where(detail => detail.AssetId > 0)
            .GroupBy(detail => detail.AssetId)
            .ToDictionary(group => group.Key, group => group.First());
        var byHostName = productDetails
            .Where(detail => !string.IsNullOrWhiteSpace(detail.HostName))
            .GroupBy(detail => detail.HostName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First());

        var merged = new List<PatchAssetDetail>();
        foreach (var host in hosts)
        {
            PatchAssetDetail? match = null;
            if (host.AgentId > 0 && byAgentId.TryGetValue(host.AgentId, out match)) { }
            else if (host.AssetId > 0 && byAssetId.TryGetValue(host.AssetId, out match)) { }
            else if (!string.IsNullOrWhiteSpace(host.HostName) &&
                     byHostName.TryGetValue(host.HostName, out match)) { }

            if (match is null)
            {
                merged.Add(host);
                continue;
            }

            merged.Add(host with
            {
                Versions = match.Versions.Count > 0 ? match.Versions : host.Versions,
                Paths = match.Paths.Count > 0 ? match.Paths : host.Paths
            });
        }

        return merged;
    }

    public async Task<PatchOperationResult> PatchApplicationsNowAsync(
        ApplicationPatchRequest request,
        CancellationToken ct = default)
    {
        ValidatePatchRequest(request);
        await EnsureApplicationPatchTargetsAsync(request, ct);
        ValidateApplicationPatchVersions(request);
        var body = BuildPatchPayload(request, ConnectSecurePatchWhen.Now, scheduledAt: null);
        return await InvokePatchAsync(request, "Application Patch", body, ct);
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

    public async Task<PatchJobListResult> GetCompanyJobsAsync(
        int companyId,
        PatchJobListQuery query,
        bool fetchRemoteJobs = true,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var page = Math.Max(1, query.Page);
        var pageSize = Math.Clamp(query.PageSize, 5, 100);
        var since = query.Since;

        var localEntries = patchActivityHistory.GetEntries(companyId, limit: 200).ToList();
        if (since is not null)
        {
            localEntries = localEntries
                .Where(entry => entry.RequestedAt >= since.Value)
                .ToList();
        }

        var remoteJobs = fetchRemoteJobs
            ? await FetchConnectSecurePatchJobsAsync(companyId, limit: 100, ct)
            : Array.Empty<PatchJobCorrelationHelper.ParsedConnectSecureJob>();

        if (since is not null)
        {
            remoteJobs = remoteJobs
                .Where(job => job.Updated is null || job.Updated.Value >= since.Value)
                .ToList();
        }

        if (fetchRemoteJobs)
            SyncLocalEntriesWithConnectSecureJobs(companyId, localEntries, remoteJobs);

        localEntries = patchActivityHistory.GetEntries(companyId, limit: 200).ToList();
        if (since is not null)
        {
            localEntries = localEntries
                .Where(entry => entry.RequestedAt >= since.Value)
                .ToList();
        }

        var linkedRemoteIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var merged = new List<PatchJobEntry>();

        foreach (var entry in localEntries)
        {
            var remote = ResolveLinkedRemoteJob(entry, remoteJobs, linkedRemoteIds);
            if (!string.IsNullOrWhiteSpace(remote?.JobId))
                linkedRemoteIds.Add(remote.JobId);
            merged.Add(ToPatchJobEntry(entry, remote));
        }

        if (!query.LocalOnly)
        {
            foreach (var remote in remoteJobs)
            {
                if (string.IsNullOrWhiteSpace(remote.JobId) || linkedRemoteIds.Contains(remote.JobId))
                    continue;
                if (query.PatchJobsOnly && !PatchJobCorrelationHelper.IsPatchJobType(remote.Type))
                    continue;

                merged.Add(ToRemotePatchJobEntry(remote));
            }
        }

        var ordered = merged
            .OrderByDescending(job => job.Updated ?? DateTimeOffset.MinValue)
            .ToList();

        return PageResults(ordered, page, pageSize);
    }

    private static PatchJobListResult PageResults(
        IReadOnlyList<PatchJobEntry> jobs,
        int page,
        int pageSize)
    {
        var total = jobs.Count;
        var items = jobs
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .ToList();
        return new PatchJobListResult(items, total, page, pageSize);
    }

    public Task<PatchVerificationResult> VerifyPatchActivityAsync(
        int companyId,
        string jobId,
        CancellationToken ct = default) =>
        VerifyPatchActivityAsync(companyId, jobId, queueInventoryRefresh: false, ct);

    public async Task<PatchVerificationResult> VerifyPatchActivityAsync(
        int companyId,
        string jobId,
        bool queueInventoryRefresh,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (string.IsNullOrWhiteSpace(jobId))
            throw new ArgumentException("Job id is required.", nameof(jobId));

        var entry = patchActivityHistory.GetByJobId(companyId, jobId);
        if (entry is null)
            throw new InvalidOperationException("Patch activity entry not found.");

        remediationService.Invalidate(companyId);
        cache.InvalidatePatchHosts(companyId);

        var remoteJob = await RefreshConnectSecureJobForEntryAsync(entry, ct);
        entry = patchActivityHistory.GetByJobId(companyId, jobId) ?? entry;

        var agentIds = entry.AgentIds?.Where(id => id > 0).Distinct().ToList() ?? [];
        string? inventoryMessage = null;
        if (queueInventoryRefresh && agentIds.Count > 0)
        {
            var scan = await scanService.TriggerAgentUpdatesAsync(companyId, agentIds, ct);
            inventoryMessage = scan.Message;
        }
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
            hostViews = await BuildOsPatchHostViewsAsync(companyId, targetFix, details, ct);
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

            var details = await GetPatchingAssetDetailsForProductAsync(
                companyId,
                solutionIds,
                entry.Product ?? "",
                ct);
            hostViews = PatchCatalogHelper.BuildHostViews(details, targetFix, isEndOfLife);
        }

        var insight = PatchCatalogHelper.BuildConnectSecureJobInsight(entry, remoteJob);
        var result = PatchCatalogHelper.BuildVerificationResult(jobId, agentIds, hostViews);
        if (PatchCatalogHelper.ShouldInferRemediationCleared(remoteJob, agentIds.Count, result.TotalHosts))
        {
            result = PatchCatalogHelper.BuildRemediationClearedVerificationResult(jobId, agentIds, result.VerifiedAt) with
            {
                ConnectSecureInsight = insight,
                InventoryRefreshMessage = inventoryMessage
            };
        }
        else
        {
            result = result with
            {
                ConnectSecureInsight = insight,
                InventoryRefreshMessage = inventoryMessage
            };
        }

        UpdateVerificationEntry(entry, result);
        return result;
    }

    private async Task<PatchJobCorrelationHelper.ParsedConnectSecureJob?> RefreshConnectSecureJobForEntryAsync(
        PatchActivityEntry entry,
        CancellationToken ct)
    {
        var remoteJobs = await FetchConnectSecurePatchJobsAsync(entry.CompanyId, limit: 100, ct);
        PatchJobCorrelationHelper.ParsedConnectSecureJob? remote = null;

        if (!string.IsNullOrWhiteSpace(entry.ConnectSecureJobId))
        {
            remote = remoteJobs.FirstOrDefault(job =>
                string.Equals(job.JobId, entry.ConnectSecureJobId, StringComparison.OrdinalIgnoreCase));
        }

        remote ??= PatchJobCorrelationHelper.FindBestMatch(entry, remoteJobs, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        if (remote is null || string.IsNullOrWhiteSpace(remote.JobId))
            return null;

        if (!string.Equals(entry.ConnectSecureJobId, remote.JobId, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(entry.ConnectSecureJobStatus, remote.Status, StringComparison.OrdinalIgnoreCase))
        {
            patchActivityHistory.UpdateEntry(entry with
            {
                ConnectSecureJobId = remote.JobId,
                ConnectSecureJobStatus = remote.Status
            });
        }

        return remote;
    }

    public async Task<IReadOnlyList<OsPendingPatchEntry>> GetOsPendingPatchesAsync(
        int companyId,
        int lookbackDays = 90,
        bool forceRefresh = false,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        // lookbackDays preserved for signature compatibility; get_remediation does not
        // restrict by discovery date — all currently-needed OS updates are returned.
        _ = lookbackDays;

        if (forceRefresh)
        {
            remediationService.Invalidate(companyId);
            cache.InvalidatePatchHosts(companyId);
        }

        var dataset = await remediationService.GetRemediationDatasetAsync(companyId, forceRefresh: forceRefresh, ct: ct);

        return dataset.Records
            .Where(r => r.IsOsUpdate)
            .Where(r => !string.IsNullOrWhiteSpace(r.Fix))
            .GroupBy(r => new { OsName = NormalizeOsName(r.Product), Fix = r.Fix.Trim() })
            .Select(g => new OsPendingPatchEntry(
                g.Key.OsName,
                ResolveOsVersion(g),
                g.Key.Fix,
                g.Select(r => r.AssetId).Where(id => id > 0).Distinct().Count(),
                g.Select(r => r.AssetId).Where(id => id > 0).Distinct().ToList()))
            .Where(entry => entry.AffectedAssets > 0)
            .OrderByDescending(entry => entry.AffectedAssets)
            .ThenBy(entry => entry.OsName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizeOsName(string product) =>
        string.IsNullOrWhiteSpace(product) ? "" : product.Trim();

    private static string ResolveOsVersion(IEnumerable<RemediationRecord> records)
    {
        // get_remediation does not return os_version directly; product name carries
        // it (e.g. "Windows 1125H2"). Caller renders OsName which already contains it.
        return "";
    }

    public async Task<IReadOnlyList<PatchHostView>> BuildOsPatchHostViewsAsync(
        int companyId,
        string targetFix,
        IEnumerable<PatchAssetDetail> details,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var detailList = details.ToList();
        if (!PatchCatalogHelper.IsOsUpdateTarget(targetFix))
            return PatchCatalogHelper.BuildHostViews(detailList, targetFix, isEndOfLife: false);

        var pendingPatches = await GetOsPendingPatchesAsync(companyId, ct: ct);
        var pendingAssetIds = PatchCatalogHelper.BuildOsPendingAssetIndex(pendingPatches, targetFix);
        return PatchCatalogHelper.BuildOsHostViews(detailList, targetFix, pendingAssetIds);
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

        var dataset = await remediationService.GetRemediationDatasetAsync(companyId, ct: ct);

        return dataset.Records
            .Where(r => r.SolutionId > 0 && !string.IsNullOrWhiteSpace(r.Product))
            .GroupBy(r => r.SolutionId)
            .Select(g =>
            {
                var primary = g.OrderByDescending(r => PatchCatalogHelper.SeverityRank(r.Severity)).First();
                return new SuppressibleRemediationEntry(
                    primary.SolutionId,
                    primary.Product,
                    primary.Fix,
                    primary.Severity,
                    primary.RemediationAction,
                    primary.IsSoftwarePatch,
                    g.Select(r => r.AssetId).Where(id => id > 0).Distinct().Count());
            })
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

        var assetWiseTask = CollectSuppressibleProblemsAsync(
            companyId,
            "/r/report_queries/asset_wise_vulnerabilities",
            AssetWiseQuery($"company_id={companyId}"),
            ct);

        var registryTask = CollectSuppressibleProblemsAsync(
            companyId,
            "/r/report_queries/registry_problems_remediation",
            ReportQuery($"company_id={companyId} and is_suppressed=false and is_remediated = false"),
            ct);

        var networkTask = CollectSuppressibleProblemsAsync(
            companyId,
            "/r/report_queries/application_vulnerabilities_net",
            ReportQuery($"company_id={companyId} and software_type='networksoftware' and unconfirmed = 'false'"),
            ct);

        var batches = await Task.WhenAll(assetWiseTask, registryTask, networkTask);

        var merged = new Dictionary<string, SuppressibleProblemEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var batch in batches)
        {
            foreach (var entry in batch)
                ConnectSecureSuppressLookup.MergeProblem(merged, entry);
        }

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

    public async Task<IReadOnlyList<PatchAssetDetail>> GetPatchingAssetDetailsForProductAsync(
        int companyId,
        IReadOnlyList<int> solutionIds,
        string productName,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");
        if (solutionIds.Count == 0)
            return [];

        var normalizedProduct = productName.Trim();
        if (!string.IsNullOrWhiteSpace(normalizedProduct) &&
            cache.TryGetPatchProductHosts(companyId, normalizedProduct, out var cachedProductHosts))
        {
            return cachedProductHosts;
        }

        var solutionSet = solutionIds.Where(id => id > 0).ToHashSet();
        var dataset = await remediationService.GetRemediationDatasetAsync(companyId, ct: ct);
        var records = dataset.Records
            .Where(r => r.IsSoftwarePatch && solutionSet.Contains(r.SolutionId))
            .ToList();

        var merged = PatchCatalogHelper.MergeAssetDetails(DeduplicateRemediationHosts(records)).ToList();
        if (!string.IsNullOrWhiteSpace(normalizedProduct))
        {
            var assetDetails = await GetOrFetchProductRemediationAssetDetailsAsync(companyId, normalizedProduct, ct);
            merged = MergeInstalledVersions(merged, assetDetails, normalizedProduct);
        }

        var enriched = await EnrichHostsWithAgentRegistryAsync(companyId, merged, ct);
        if (!string.IsNullOrWhiteSpace(normalizedProduct))
            cache.SetPatchProductHosts(companyId, normalizedProduct, enriched);

        return enriched;
    }

    private static List<PatchAssetDetail> DeduplicateRemediationHosts(IEnumerable<RemediationRecord> records)
    {
        var deduped = new Dictionary<string, PatchAssetDetail>(StringComparer.OrdinalIgnoreCase);
        foreach (var record in records)
        {
            var key = $"{record.AgentId}:{record.AssetId}:{record.HostName}";
            if (string.IsNullOrEmpty(key.Trim(':')))
                continue;

            if (!deduped.TryGetValue(key, out var existing) ||
                (!existing.OnlineStatus && record.OnlineStatus))
            {
                deduped[key] = BuildPatchAssetDetailFromRecord(record);
            }
        }

        return deduped.Values.ToList();
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
        await TryLinkConnectSecureJobAsync(request.CompanyId, jobId, ct);
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
            request.OsAssetIds.Where(id => id > 0).Distinct().ToList(),
            AssetIds: request.AssetIds.Where(id => id > 0).Distinct().ToList()));

        return jobId;
    }

    private void UpdateVerificationEntry(PatchActivityEntry entry, PatchVerificationResult result)
    {
        patchActivityHistory.UpdateEntry(entry with
        {
            VersionCheckStatus = result.Status,
            VerificationSummary = result.Summary,
            VerifiedAt = result.VerifiedAt
        });
    }

    private async Task TryLinkConnectSecureJobAsync(int companyId, string localJobId, CancellationToken ct)
    {
        var entry = patchActivityHistory.GetByJobId(companyId, localJobId);
        if (entry is null)
            return;

        for (var attempt = 0; attempt < 4; attempt++)
        {
            if (attempt > 0)
                await Task.Delay(TimeSpan.FromSeconds(10), ct);

            var remoteJobs = await FetchConnectSecurePatchJobsAsync(companyId, limit: 50, ct);
            var linkedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var match = PatchJobCorrelationHelper.FindBestMatch(entry, remoteJobs, linkedIds);
            if (match is null || string.IsNullOrWhiteSpace(match.JobId))
                continue;

            patchActivityHistory.UpdateEntry(entry with
            {
                ConnectSecureJobId = match.JobId,
                ConnectSecureJobStatus = match.Status
            });
            return;
        }
    }

    private void SyncLocalEntriesWithConnectSecureJobs(
        int companyId,
        IReadOnlyList<PatchActivityEntry> localEntries,
        IReadOnlyList<PatchJobCorrelationHelper.ParsedConnectSecureJob> remoteJobs)
    {
        var linkedRemoteIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in localEntries)
        {
            var working = ClearStaleConnectSecureJobLink(entry, remoteJobs);
            var remote = PatchJobCorrelationHelper.FindBestMatch(working, remoteJobs, linkedRemoteIds);
            if (remote is null || string.IsNullOrWhiteSpace(remote.JobId))
                continue;

            linkedRemoteIds.Add(remote.JobId);
            var needsUpdate = !string.Equals(working.ConnectSecureJobId, remote.JobId, StringComparison.OrdinalIgnoreCase) ||
                              !string.Equals(working.ConnectSecureJobStatus, remote.Status, StringComparison.OrdinalIgnoreCase);
            if (!needsUpdate)
                continue;

            patchActivityHistory.UpdateEntry(working with
            {
                ConnectSecureJobId = remote.JobId,
                ConnectSecureJobStatus = remote.Status
            });
        }
    }

    private PatchActivityEntry ClearStaleConnectSecureJobLink(
        PatchActivityEntry entry,
        IReadOnlyList<PatchJobCorrelationHelper.ParsedConnectSecureJob> remoteJobs)
    {
        if (string.IsNullOrWhiteSpace(entry.ConnectSecureJobId))
            return entry;

        var linked = remoteJobs.FirstOrDefault(job =>
            string.Equals(job.JobId, entry.ConnectSecureJobId, StringComparison.OrdinalIgnoreCase));
        if (linked is null || PatchJobCorrelationHelper.ProductNameMatches(entry, linked))
            return entry;

        var cleared = entry with
        {
            ConnectSecureJobId = null,
            ConnectSecureJobStatus = null
        };
        patchActivityHistory.UpdateEntry(cleared);
        return cleared;
    }

    private static PatchJobCorrelationHelper.ParsedConnectSecureJob? ResolveLinkedRemoteJob(
        PatchActivityEntry entry,
        IReadOnlyList<PatchJobCorrelationHelper.ParsedConnectSecureJob> remoteJobs,
        IReadOnlySet<string> alreadyLinkedRemoteIds)
    {
        if (!string.IsNullOrWhiteSpace(entry.ConnectSecureJobId))
        {
            var byId = remoteJobs.FirstOrDefault(job =>
                string.Equals(job.JobId, entry.ConnectSecureJobId, StringComparison.OrdinalIgnoreCase));
            if (byId is not null && PatchJobCorrelationHelper.ProductNameMatches(entry, byId))
                return byId;
        }

        return PatchJobCorrelationHelper.FindBestMatch(entry, remoteJobs, alreadyLinkedRemoteIds);
    }

    private async Task<IReadOnlyList<PatchJobCorrelationHelper.ParsedConnectSecureJob>> FetchConnectSecurePatchJobsAsync(
        int companyId,
        int limit,
        CancellationToken ct)
    {
        var rows = await ConnectSecurePagedQuery.FetchCompanyScopedPagesByIndexAsync(
            (query, token) => FetchArrayAsync("/r/report_queries/patch_jobview", query, token),
            new Dictionary<string, string>
            {
                ["condition"] = $"company_id={companyId}",
                ["order_by"] = "created desc"
            },
            companyId,
            ct,
            pageSize: 100,
            maxPages: 3);

        return rows
            .Select(PatchJobCorrelationHelper.ParsePatchJobViewRow)
            .Where(job => !string.IsNullOrWhiteSpace(job.JobId))
            .OrderByDescending(job => job.Updated ?? DateTimeOffset.MinValue)
            .Take(limit)
            .ToList();
    }

    private static PatchJobEntry ToPatchJobEntry(PatchActivityEntry entry, PatchJobCorrelationHelper.ParsedConnectSecureJob? remote)
    {
        var canVerify = entry.AgentIds is { Count: > 0 } &&
                        (entry.SolutionIds is { Count: > 0 } || entry.IsOsPatch);

        var description = !string.IsNullOrWhiteSpace(remote?.Description)
            ? remote.Description
            : entry.Description;

        var csJobId = entry.ConnectSecureJobId ?? remote?.JobId;

        var versionCheck = PatchJobCorrelationHelper.ResolveVersionCheckStatus(entry);

        var jobStatus = entry.ConnectSecureJobStatus ?? remote?.Status;
        if (string.IsNullOrWhiteSpace(jobStatus))
            jobStatus = string.IsNullOrWhiteSpace(csJobId) ? "Accepted" : "Submitted";

        return new PatchJobEntry(
            entry.JobId,
            entry.Type,
            PatchJobCorrelationHelper.FormatJobStatusLabel(jobStatus),
            description,
            entry.HostName,
            entry.AgentIp,
            remote?.Updated ?? entry.VerifiedAt ?? entry.RequestedAt,
            IsLocal: true,
            CanVerify: canVerify,
            VerificationSummary: entry.VerificationSummary,
            ConnectSecureJobId: csJobId,
            VersionCheckStatus: string.IsNullOrWhiteSpace(versionCheck) ? null : versionCheck,
            AgentId: remote?.AgentId ?? entry.AgentIds?.FirstOrDefault(),
            TargetFix: entry.TargetFix,
            Product: entry.Product);
    }

    private static PatchJobEntry ToRemotePatchJobEntry(PatchJobCorrelationHelper.ParsedConnectSecureJob remote) =>
        new(
            remote.JobId,
            remote.Type,
            PatchJobCorrelationHelper.FormatJobStatusLabel(remote.Status),
            remote.Description,
            remote.HostName,
            remote.AgentIp,
            remote.Updated,
            IsLocal: false,
            CanVerify: false,
            ConnectSecureJobId: remote.JobId,
            AgentId: remote.AgentId,
            Product: remote.ProductName);

    private static void ValidatePatchRequest(ApplicationPatchRequest request)
    {
        if (request.CompanyId <= 0)
            throw new InvalidOperationException("Company id is required.");
        if (request.PatchType != ConnectSecurePatchType.Os && request.IncludedApplications.Count == 0)
            throw new InvalidOperationException("At least one application must be included in the patch.");
        if (request.AssetIds.Count == 0 && request.AgentIds.Count == 0)
            throw new InvalidOperationException("Select at least one asset or agent to patch.");
    }

    public static void ValidateApplicationPatchVersions(ApplicationPatchRequest request)
    {
        if (request.PatchType != ConnectSecurePatchType.App)
            return;

        if (!PatchCatalogHelper.IsVersionComparableTarget(request.TargetFix))
            return;

        if (request.FromVersions.Count == 0)
        {
            throw new InvalidOperationException(
                "ConnectSecure requires the current installed version (from_versions) to queue this patch. " +
                "Expand the product, click Refresh hosts, confirm Installed shows a version, then try again.");
        }

        var assetIds = request.AssetIds.Where(id => id > 0).ToHashSet();
        var coveredAssets = request.FromVersions.Keys
            .Select(key => int.TryParse(key, out var id) ? id : 0)
            .Where(id => id > 0)
            .ToHashSet();

        if (assetIds.Count > 0 && coveredAssets.Count < assetIds.Count)
        {
            throw new InvalidOperationException(
                "Installed version is missing for one or more selected hosts. Refresh hosts before patching.");
        }
    }

    private async Task EnsureApplicationPatchTargetsAsync(
        ApplicationPatchRequest request,
        CancellationToken ct)
    {
        if (request.PatchType != ConnectSecurePatchType.App)
            return;

        var agentSet = request.AgentIds.Where(id => id > 0).ToHashSet();
        if (agentSet.Count == 0)
            return;

        var dataset = await remediationService.GetRemediationDatasetAsync(request.CompanyId, ct: ct);
        var records = dataset.Records
            .Where(r => r.IsSoftwarePatch && agentSet.Contains(r.AgentId))
            .Where(r => request.IncludedApplications.Any(product =>
                PatchCatalogHelper.ProductNamesMatch(r.Product, product)))
            .ToList();

        if (records.Count == 0)
            return;

        var assetIds = request.AssetIds.Where(id => id > 0).Distinct().ToList();
        foreach (var record in records)
        {
            if (record.AssetId > 0 && !assetIds.Contains(record.AssetId))
                assetIds.Add(record.AssetId);
        }

        if (assetIds.Count == 0)
            assetIds.AddRange(agentSet);

        request.AssetIds = assetIds;

        if (request.FromVersions.Count == 0)
            return;

        var normalized = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var record in records)
        {
            if (record.AssetId <= 0)
                continue;

            var agentKey = record.AgentId.ToString();
            if (request.FromVersions.TryGetValue(agentKey, out var byAgent))
            {
                normalized[record.AssetId.ToString()] = byAgent;
                continue;
            }

            var assetKey = record.AssetId.ToString();
            if (request.FromVersions.TryGetValue(assetKey, out var byAsset))
                normalized[assetKey] = byAsset;
        }

        foreach (var pair in request.FromVersions)
        {
            normalized.TryAdd(pair.Key, pair.Value);
        }

        request.FromVersions = normalized;
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

    public void InvalidateCompanyCache(int companyId) => cache.InvalidateCompany(companyId);

    public void InvalidatePatchHostsCache(int companyId) => cache.InvalidatePatchHosts(companyId);

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

        await PopulateAgentLookupsFromCompanyPagesAsync(
            companyId,
            neededAgentIds,
            neededAssetIds,
            neededHostNames,
            byAgentId,
            byAssetId,
            byHostName,
            ct);

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

    private async Task PopulateAgentLookupsFromCompanyPagesAsync(
        int companyId,
        HashSet<int> neededAgentIds,
        HashSet<int> neededAssetIds,
        HashSet<string> neededHostNames,
        Dictionary<int, PatchAssetDetail> byAgentId,
        Dictionary<int, PatchAssetDetail> byAssetId,
        Dictionary<string, PatchAssetDetail> byHostName,
        CancellationToken ct)
    {
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

    private static PatchableApplicationEntry BuildPatchableApplicationEntry(IGrouping<int, RemediationRecord> group)
    {
        var rows = group.ToList();
        var primary = rows
            .OrderByDescending(r => PatchCatalogHelper.SeverityRank(r.Severity))
            .ThenByDescending(r => r.TotalVulsCount)
            .First();

        var assetIds = rows
            .Select(r => r.AssetId)
            .Where(id => id > 0)
            .Distinct()
            .ToList();

        return new PatchableApplicationEntry(
            primary.SolutionId,
            primary.Product,
            primary.Fix,
            primary.IsSoftwarePatch,
            assetIds.Count,
            assetIds,
            primary.Severity,
            primary.RemediationAction);
    }

    private static PatchAssetDetail BuildPatchAssetDetailFromRecord(RemediationRecord record) =>
        new(
            AssetId: record.AssetId,
            Ip: record.Ip,
            HostName: record.HostName,
            AgentId: record.AgentId,
            OnlineStatus: record.OnlineStatus,
            ApplicationNames: string.IsNullOrWhiteSpace(record.Product) ? [] : new[] { record.Product },
            Versions: [],
            Paths: []);

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

        if (value.ValueKind == JsonValueKind.Object)
        {
            var version = ConnectSecureJsonReader.GetString(value, "version");
            return PatchCatalogHelper.SplitVersionTokens(version);
        }

        if (value.ValueKind == JsonValueKind.String)
            return PatchCatalogHelper.SplitVersionTokens(value.GetString());

        if (value.ValueKind != JsonValueKind.Array)
            return [];

        var versions = new List<string>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Object)
            {
                versions.AddRange(PatchCatalogHelper.SplitVersionTokens(
                    ConnectSecureJsonReader.GetString(item, "version")));
            }
            else if (item.ValueKind == JsonValueKind.String)
            {
                versions.AddRange(PatchCatalogHelper.SplitVersionTokens(item.GetString()));
            }
        }

        return PatchCatalogHelper.NormalizeVersions(versions);
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

    private async Task<List<JsonElement>> FetchArrayAsync(
        string endpoint,
        IReadOnlyDictionary<string, string> query,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, endpoint, query, ct: ct);
        return ConnectSecureJsonReader.ExtractDataArray(response);
    }

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

    private async Task<List<SuppressibleProblemEntry>> CollectSuppressibleProblemsAsync(
        int companyId,
        string endpoint,
        Dictionary<string, string> query,
        CancellationToken ct)
    {
        var rows = await FetchAllPagesAsync(endpoint, query, ct);
        var results = new List<SuppressibleProblemEntry>();
        foreach (var row in rows)
        {
            if (!MatchesCompanyInIds(row, companyId))
                continue;

            results.Add(ParseSuppressibleProblem(row));
        }

        return results;
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
