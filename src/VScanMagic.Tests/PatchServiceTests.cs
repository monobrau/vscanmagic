using System.Text.Json;
using VScanMagic.Core.Services;
using VScanMagic.ConnectSecure;

namespace VScanMagic.Tests;

public sealed class ConnectSecurePatchServiceTests
{
    [Fact]
    public void ValidateApplicationPatchVersions_RequiresFromVersionsForSemverTargets()
    {
        var request = new ApplicationPatchRequest
        {
            CompanyId = 1,
            AssetIds = [24079295],
            AgentIds = [402780],
            IncludedApplications = ["Mozilla Firefox"],
            TargetFix = "151.0.0",
            FromVersions = []
        };

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ConnectSecurePatchService.ValidateApplicationPatchVersions(request));
        Assert.Contains("from_versions", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateApplicationPatchVersions_AllowsQualitativeTargetsWithoutVersions()
    {
        var request = new ApplicationPatchRequest
        {
            CompanyId = 1,
            AssetIds = [1],
            AgentIds = [2],
            IncludedApplications = ["Microsoft 365"],
            TargetFix = "Latest Patch",
            FromVersions = []
        };

        ConnectSecurePatchService.ValidateApplicationPatchVersions(request);
    }

    [Fact]
    public void BuildPatchPayload_PatchNow_IncludesRequiredFields()
    {
        var request = new ApplicationPatchRequest
        {
            CompanyId = 42,
            AssetIds = [6],
            AgentIds = [6],
            IncludedApplications = ["7-Zip"],
            FromVersions = new Dictionary<string, string> { ["6"] = "24.09" }
        };

        var body = ConnectSecurePatchService.BuildPatchPayload(request, ConnectSecurePatchWhen.Now, scheduledAt: null);

        Assert.Equal("now", body["patch_when"]);
        Assert.Equal(new[] { 42 }, body["companies"]);
        Assert.Equal(new[] { 6 }, body["assets"]);
        Assert.Equal(new[] { 6 }, body["agents_id"]);
        Assert.Equal(new[] { "7-Zip" }, body["included_application"]);
        Assert.Equal("application", body["type"]);
        Assert.False(body.ContainsKey("patch_type"));
        Assert.False(body.ContainsKey("date"));
    }

    [Fact]
    public void BuildPatchPayload_PatchLater_IncludesScheduleFields()
    {
        var request = new ScheduledApplicationPatchRequest
        {
            CompanyId = 1,
            AssetIds = [1],
            IncludedApplications = ["Bitwarden"],
            FromVersions = new Dictionary<string, string> { ["1"] = "2025.7.0" },
            PatchType = ConnectSecurePatchType.App
        };

        var scheduledAt = new DateTime(2025, 10, 24, 19, 0, 0);
        var body = ConnectSecurePatchService.BuildPatchPayload(request, ConnectSecurePatchWhen.Later, scheduledAt);

        Assert.Equal("later", body["patch_when"]);
        Assert.Equal("app", body["patch_type"]);
        Assert.Equal("application", body["type"]);

        var date = Assert.IsType<Dictionary<string, int>>(body["date"]);
        Assert.Equal(24, date["days"]);
        Assert.Equal(10, date["months"]);
        Assert.Equal(2025, date["years"]);

        var time = Assert.IsType<Dictionary<string, int>>(body["time"]);
        Assert.Equal(19, time["hours"]);
        Assert.Equal(0, time["minutes"]);
    }

    [Fact]
    public void BuildPatchPayload_OsPatchNow_IncludesPatchTypeAndReboot()
    {
        var request = new ApplicationPatchRequest
        {
            CompanyId = 8604,
            AssetIds = [70552, 97029],
            IncludedApplications = ["Windows 1022H2"],
            PatchType = ConnectSecurePatchType.Os,
            TriggerReboot = true
        };

        var body = ConnectSecurePatchService.BuildPatchPayload(request, ConnectSecurePatchWhen.Now, scheduledAt: null);

        Assert.Equal("now", body["patch_when"]);
        Assert.Equal("os", body["patch_type"]);
        Assert.Equal(true, body["trigger_reboot"]);
        Assert.False(body.ContainsKey("type"));
    }

    [Fact]
    public void Clone_DoesNotMutateOriginalPatchType()
    {
        var request = new ApplicationPatchRequest
        {
            CompanyId = 1,
            AgentIds = [1],
            IncludedApplications = ["App"],
            PatchType = ConnectSecurePatchType.App
        };

        var copy = request.Clone();
        copy.PatchType = ConnectSecurePatchType.Os;

        Assert.Equal(ConnectSecurePatchType.App, request.PatchType);
        Assert.Equal(ConnectSecurePatchType.Os, copy.PatchType);
    }
}

public sealed class PatchCatalogHelperTests
{
    [Fact]
    public void GroupByProduct_MergesDuplicateProductRows()
    {
        var grouped = PatchCatalogHelper.GroupByProduct(
        [
            new PatchableApplicationEntry(1, "Google Chrome", "148.0.1", true, 5, [], "Critical", "Software Patch"),
            new PatchableApplicationEntry(2, "Google Chrome", "148.0.1", true, 3, [], "Critical", "Software Patch"),
            new PatchableApplicationEntry(3, "TeamViewer", "15.76.6", true, 4, [], "High", "Software Patch")
        ]);

        Assert.Equal(2, grouped.Count);
        var chrome = grouped.First(group => group.Product == "Google Chrome");
        Assert.Equal(2, chrome.RemediationEntryCount);
        Assert.Equal(5, chrome.ReportedAssets);
        Assert.Equal(2, chrome.SolutionIds.Count);
    }

    [Fact]
    public void MergeAssetDetails_PrefersOnlineWhenSameHostAppearsTwice()
    {
        var merged = PatchCatalogHelper.MergeAssetDetails(
        [
            new PatchAssetDetail(1, "10.0.0.5", "host-a", 10, false, ["App"], ["1.0"], []),
            new PatchAssetDetail(1, "10.0.0.5", "host-a", 10, true, ["App"], ["1.0"], [])
        ]);

        Assert.Single(merged);
        Assert.True(merged[0].OnlineStatus);
    }

    [Fact]
    public void MergeAssetDetails_DoesNotCollideAssetIdWithAgentId()
    {
        var merged = PatchCatalogHelper.MergeAssetDetails(
        [
            new PatchAssetDetail(5, "10.0.0.1", "asset-host", 100, true, ["App"], ["1.0"], []),
            new PatchAssetDetail(0, "10.0.0.2", "agent-host", 5, true, ["App"], ["2.0"], [])
        ]);

        Assert.Equal(2, merged.Count);
        Assert.Contains(merged, detail => detail.HostName == "asset-host");
        Assert.Contains(merged, detail => detail.HostName == "agent-host");
    }

    [Fact]
    public void FormatVersionSummary_DeduplicatesRepeatedVersions()
    {
        var summary = PatchCatalogHelper.FormatVersionSummary(["144.0.2", "144.0.2", "147.0.1"]);
        Assert.Equal("144.0.2, 147.0.1", summary);
    }

    [Fact]
    public void DetermineHostStatus_MarksQualitativeTargetAsPendingWithoutInstalledVersion()
    {
        var status = PatchCatalogHelper.DetermineHostStatus(
            online: true,
            isEndOfLife: false,
            installedVersions: [],
            targetFix: "Latest Patch");

        Assert.Equal(HostPatchStatus.Pending, status);
    }

    [Fact]
    public void DetermineHostStatus_MarksUnknownWhenVersionTargetHasNoInstalledData()
    {
        var status = PatchCatalogHelper.DetermineHostStatus(
            online: true,
            isEndOfLife: false,
            installedVersions: [],
            targetFix: "148.0.7778.179");

        Assert.Equal(HostPatchStatus.Unknown, status);
    }

    [Fact]
    public void IsVersionComparableTarget_DetectsSemverAndKbFixes()
    {
        Assert.True(PatchCatalogHelper.IsVersionComparableTarget("148.0.7778.179"));
        Assert.True(PatchCatalogHelper.IsVersionComparableTarget("KB5034765"));
        Assert.False(PatchCatalogHelper.IsVersionComparableTarget("Latest Patch"));
        Assert.False(PatchCatalogHelper.IsVersionComparableTarget("Upgrade to supported Operating System is recommended"));
    }

    [Fact]
    public void DetermineHostStatus_MarksOfflineBeforePending()
    {
        var status = PatchCatalogHelper.DetermineHostStatus(
            online: false,
            isEndOfLife: false,
            installedVersions: ["1.0"],
            targetFix: "2.0");

        Assert.Equal(HostPatchStatus.Offline, status);
    }

    [Fact]
    public void DetermineHostStatus_MarksAtTargetWhenInstalledMatchesFix()
    {
        var status = PatchCatalogHelper.DetermineHostStatus(
            online: true,
            isEndOfLife: false,
            installedVersions: ["148.0.7778.179"],
            targetFix: "148.0.7778.179");

        Assert.Equal(HostPatchStatus.AtTarget, status);
    }

    [Fact]
    public void DetermineHostStatus_MarksAtTargetWhenInstalledMatchesFixPrefix()
    {
        var status = PatchCatalogHelper.DetermineHostStatus(
            online: true,
            isEndOfLife: false,
            installedVersions: ["3.2.4.0"],
            targetFix: "3.2.4");

        Assert.Equal(HostPatchStatus.AtTarget, status);
    }

    [Fact]
    public void DetermineOsHostStatus_MarksAtTargetWhenKbNoLongerPending()
    {
        var detail = new PatchAssetDetail(24076444, "10.0.0.1", "Zillah3", 402714, true, ["Windows 11"], ["10.0.26200"], []);
        var pendingAssets = new HashSet<int> { 99999 };

        var status = PatchCatalogHelper.DetermineOsHostStatus(
            online: true,
            detail,
            targetFix: "5089549",
            pendingAssets);

        Assert.Equal(HostPatchStatus.AtTarget, status);
    }

    [Fact]
    public void DetermineOsHostStatus_MarksPendingWhenKbStillPending()
    {
        var detail = new PatchAssetDetail(24076444, "10.0.0.1", "Zillah3", 402714, true, ["Windows 11"], ["10.0.26200"], []);
        var pendingAssets = new HashSet<int> { 24076444 };

        var status = PatchCatalogHelper.DetermineOsHostStatus(
            online: true,
            detail,
            targetFix: "KB5089549",
            pendingAssets);

        Assert.Equal(HostPatchStatus.Pending, status);
    }

    [Fact]
    public void BuildOsPendingAssetIndex_MatchesKbWithOrWithoutPrefix()
    {
        var pending = new[]
        {
            new OsPendingPatchEntry("Windows 11", "10.0.26200", "5089549", 1, [24076444, 99999])
        };

        var index = PatchCatalogHelper.BuildOsPendingAssetIndex(pending, "KB5089549");

        Assert.Contains(24076444, index);
        Assert.Contains(99999, index);
    }

    [Fact]
    public void BuildHostViews_UsesTargetFixForComparison()
    {
        var views = PatchCatalogHelper.BuildHostViews(
        [
            new PatchAssetDetail(1, "10.0.0.2", "host-a", 10, true, ["Chrome"], ["148.0.1"], [])
        ],
            "148.0.1",
            isEndOfLife: false);

        Assert.Single(views);
        Assert.Equal(HostPatchStatus.AtTarget, views[0].Status);
    }

    [Fact]
    public void BuildVerificationResult_MarksVerifiedWhenAllOnlineHostsAtTarget()
    {
        var result = PatchCatalogHelper.BuildVerificationResult(
            "job-1",
            [10, 11],
            [
                new PatchHostView(
                    new PatchAssetDetail(1, "10.0.0.1", "host-a", 10, true, ["App"], ["2.0"], []),
                    HostPatchStatus.AtTarget,
                    "At target"),
                new PatchHostView(
                    new PatchAssetDetail(2, "10.0.0.2", "host-b", 11, true, ["App"], ["2.0"], []),
                    HostPatchStatus.AtTarget,
                    "At target")
            ]);

        Assert.Equal("Verified", result.Status);
        Assert.Equal(2, result.AtTargetCount);
        Assert.Equal(2, result.TotalHosts);
        Assert.Contains("2/2 at target", result.Summary);
    }

    [Fact]
    public void BuildVerificationResult_MarksPartialWhenSomeHostsPending()
    {
        var result = PatchCatalogHelper.BuildVerificationResult(
            "job-2",
            [10, 11],
            [
                new PatchHostView(
                    new PatchAssetDetail(1, "10.0.0.1", "host-a", 10, true, ["App"], ["2.0"], []),
                    HostPatchStatus.AtTarget,
                    "At target"),
                new PatchHostView(
                    new PatchAssetDetail(2, "10.0.0.2", "host-b", 11, true, ["App"], ["1.0"], []),
                    HostPatchStatus.Pending,
                    "Pending patch")
            ]);

        Assert.Equal("Partial", result.Status);
        Assert.Equal(1, result.AtTargetCount);
        Assert.Equal(1, result.PendingCount);
    }

    [Fact]
    public void BuildVerificationResult_ExcludesOfflineFromVerifiedCount()
    {
        var result = PatchCatalogHelper.BuildVerificationResult(
            "job-3",
            [10, 11],
            [
                new PatchHostView(
                    new PatchAssetDetail(1, "10.0.0.1", "host-a", 10, true, ["App"], ["2.0"], []),
                    HostPatchStatus.AtTarget,
                    "At target"),
                new PatchHostView(
                    new PatchAssetDetail(2, "10.0.0.2", "host-b", 11, false, ["App"], ["1.0"], []),
                    HostPatchStatus.Offline,
                    "Offline")
            ]);

        Assert.Equal("Verified", result.Status);
        Assert.Equal(1, result.OfflineCount);
    }
}

public sealed class PatchJobCorrelationHelperTests
{
    [Fact]
    public void FindBestMatch_LinksByAgentTimeAndProduct()
    {
        var local = new PatchActivityEntry(
            17624,
            "local-1",
            "Application Patch",
            "Submitted",
            "GIMP on Zillah3.dorks.lan",
            "Zillah3.dorks.lan",
            null,
            new DateTimeOffset(2026, 5, 25, 17, 59, 50, TimeSpan.FromHours(-5)),
            "Message sent for patch update",
            [402714],
            [47],
            "GIMP",
            "3.2.4");

        var remote = new PatchJobCorrelationHelper.ParsedConnectSecureJob(
            "cs-job-123",
            "Application",
            "Success",
            "GIMP — Success: 1, Failed: 0, Pending: 0",
            "Zillah3.dorks.lan",
            "10.0.0.1",
            402714,
            new DateTimeOffset(2026, 5, 25, 18, 0, 0, TimeSpan.FromHours(-5)),
            ProductName: "GIMP");

        var match = PatchJobCorrelationHelper.FindBestMatch(local, [remote], new HashSet<string>());
        Assert.NotNull(match);
        Assert.Equal("cs-job-123", match!.JobId);
    }

    [Fact]
    public void ParsePatchJobViewRow_ParsesPortalFirefoxJob()
    {
        const string json = """
            {
              "job_id": "ac71ebf5-fcbe-4b57-9995-ff02ae21435a",
              "job_status": "Initiated",
              "product_name": "Mozilla Firefox",
              "type": "Application",
              "created": "2026-05-26T02:23:16.827459",
              "msg": [0, 0, 2],
              "patch_job_details": {
                "24079295": {
                  "from_version": "150.0.1",
                  "host_name": "Roswell.dorks.lan",
                  "status": "Pending",
                  "to_version": "151.0.1"
                }
              }
            }
            """;

        using var doc = JsonDocument.Parse(json);
        var parsed = PatchJobCorrelationHelper.ParsePatchJobViewRow(doc.RootElement);

        Assert.Equal("ac71ebf5-fcbe-4b57-9995-ff02ae21435a", parsed.JobId);
        Assert.Equal("Initiated", parsed.Status);
        Assert.Equal("Mozilla Firefox", parsed.ProductName);
        Assert.Equal(0, parsed.SuccessCount);
        Assert.Equal(2, parsed.PendingCount);
        Assert.Contains(24079295, parsed.AssetIds!);
        Assert.Contains("Mozilla Firefox", parsed.Description);
    }

    [Fact]
    public void FindBestMatch_LinksByAssetProductAndTime()
    {
        var local = new PatchActivityEntry(
            17624,
            "local-1",
            "Application Patch",
            "Submitted",
            "Mozilla Firefox on Roswell.dorks.lan",
            "Roswell.dorks.lan",
            null,
            new DateTimeOffset(2026, 5, 25, 21, 23, 16, TimeSpan.FromHours(-5)),
            "Message sent for patch update",
            [402780],
            [47],
            "Mozilla Firefox",
            "151.0.0",
            AssetIds: [24079295]);

        var remote = PatchJobCorrelationHelper.ParsePatchJobViewRow(
            JsonDocument.Parse("""
                {
                  "job_id": "ac71ebf5-fcbe-4b57-9995-ff02ae21435a",
                  "job_status": "Initiated",
                  "product_name": "Mozilla Firefox",
                  "type": "Application",
                  "created": "2026-05-26T02:23:16.827459",
                  "msg": [0, 0, 1],
                  "patch_job_details": {
                    "24079295": { "host_name": "Roswell.dorks.lan", "status": "Pending" }
                  }
                }
                """).RootElement);

        var match = PatchJobCorrelationHelper.FindBestMatch(local, [remote], new HashSet<string>());
        Assert.NotNull(match);
        Assert.Equal("ac71ebf5-fcbe-4b57-9995-ff02ae21435a", match!.JobId);
    }

    [Fact]
    public void ResolveVersionCheckStatus_PrefersExplicitField()
    {
        var entry = new PatchActivityEntry(
            1,
            "job",
            "Application Patch",
            "Submitted",
            "desc",
            null,
            null,
            DateTimeOffset.UtcNow,
            null,
            VersionCheckStatus: "Verified");

        Assert.Equal("Verified", PatchJobCorrelationHelper.ResolveVersionCheckStatus(entry));
    }

    [Fact]
    public void PatchJobListQuery_Since_ReturnsNullForAllTime()
    {
        var query = new PatchJobListQuery(DaysBack: 0);
        Assert.Null(query.Since);
    }

    [Fact]
    public void PatchJobListQuery_Since_ReturnsCutoffForRecentWindow()
    {
        var query = new PatchJobListQuery(DaysBack: 7);
        Assert.NotNull(query.Since);
        Assert.True(query.Since!.Value <= DateTimeOffset.Now.AddDays(-6.9));
    }

    [Fact]
    public void IsInProgressJobStatus_DetectsPortalStatuses()
    {
        Assert.True(PatchJobCorrelationHelper.IsInProgressJobStatus("In Progress"));
        Assert.True(PatchJobCorrelationHelper.IsInProgressJobStatus("Pending"));
        Assert.True(PatchJobCorrelationHelper.IsInProgressJobStatus("Initiated"));
        Assert.False(PatchJobCorrelationHelper.IsTerminalJobStatus("In Progress"));
        Assert.True(PatchJobCorrelationHelper.IsTerminalJobStatus("Success"));
    }
}

public sealed class ConnectSecureJsonReaderTests
{
    [Fact]
    public void ExtractDataArray_WrapsSingleObjectInData()
    {
        using var doc = JsonDocument.Parse("""{"status":true,"data":{"solution_id":47,"product":"GIMP"}}""");
        var rows = ConnectSecureJsonReader.ExtractDataArray(doc.RootElement);
        Assert.Single(rows);
        Assert.Equal(47, ConnectSecureJsonReader.GetInt(rows[0], "solution_id"));
        Assert.Equal("GIMP", ConnectSecureJsonReader.GetString(rows[0], "product"));
    }
}

public sealed class AgentConnectivityHelperTests
{
    [Fact]
    public void IsOnlineFromLastPing_UsesLastPingNotLastReported()
    {
        var now = new DateTime(2026, 5, 24, 18, 0, 0, DateTimeKind.Utc);
        var lastPing = "2026-05-14T14:45:18.696462";
        var lastReported = "2026-05-24T13:25:08.143486";

        Assert.False(AgentConnectivityHelper.IsOnlineFromLastPing(lastPing, now));
        Assert.False(AgentConnectivityHelper.IsOnlineFromAgentTimestamps(lastPing, lastReported, now));
        Assert.False(AgentConnectivityHelper.IsOnlineFromAgentTimestamps(null, lastReported, now));
        Assert.True(AgentConnectivityHelper.IsOnlineFromAgentTimestamps(null, "2026-05-24T17:30:00", now));
    }

    [Fact]
    public void IsOnlineFromLastPing_UsesOneHourThreshold()
    {
        var now = new DateTime(2026, 5, 24, 18, 0, 0, DateTimeKind.Utc);
        Assert.True(AgentConnectivityHelper.IsOnlineFromLastPing("2026-05-24T17:30:00Z", now));
        Assert.False(AgentConnectivityHelper.IsOnlineFromLastPing("2026-05-24T15:00:00Z", now));
    }

    [Fact]
    public void FormatAgentTypeLabel_ProbeAndLightweight()
    {
        Assert.Equal("Probe", AgentConnectivityHelper.FormatAgentTypeLabel("PROBE"));
        Assert.Equal("Lightweight", AgentConnectivityHelper.FormatAgentTypeLabel("LIGHTWEIGHT"));
    }
}

public sealed class ConnectSecurePagedQueryTests
{
    [Fact]
    public async Task FetchCompanyScopedPagesAsync_StopsAfterCompanyRowsEnd()
    {
        var pages = new List<List<(int CompanyId, int Id)>>
        {
            Enumerable.Range(1, 5000).Select(i => (CompanyId: i <= 25 ? 100 : 200, Id: i)).ToList(),
            Enumerable.Range(1, 5000).Select(i => (CompanyId: i <= 57 ? 100 : 200, Id: 5000 + i)).ToList(),
            Enumerable.Range(1, 90).Select(i => (CompanyId: 200, Id: 10000 + i)).ToList(),
        };

        var call = 0;
        Task<List<JsonElement>> Fetch(Dictionary<string, string> query, CancellationToken _)
        {
            var page = call++;
            if (page >= pages.Count)
                return Task.FromResult(new List<JsonElement>());

            var json = pages[page]
                .Select(row =>
                {
                    using var doc = JsonDocument.Parse($"{{\"company_id\":{row.CompanyId},\"id\":{row.Id}}}");
                    return doc.RootElement.Clone();
                })
                .ToList();
            return Task.FromResult(json);
        }

        var rows = await ConnectSecurePagedQuery.FetchCompanyScopedPagesAsync(
            Fetch,
            new Dictionary<string, string> { ["order_by"] = "affected_assets desc" },
            companyId: 100,
            CancellationToken.None,
            pageSize: 5000,
            maxPages: 10);

        Assert.Equal(82, rows.Count);
        Assert.All(rows, row => Assert.Equal(100, ConnectSecureJsonReader.GetInt(row, "company_id")));
    }

    [Fact]
    public async Task FetchAllPagesByIndexAsync_StopsWhenSkipIgnoredAndPageRepeats()
    {
        var batch = Enumerable.Range(1, 5)
            .Select(i =>
            {
                using var doc = JsonDocument.Parse($"{{\"solution_id\":{i}}}");
                return doc.RootElement.Clone();
            })
            .ToList();

        var call = 0;
        Task<List<JsonElement>> Fetch(Dictionary<string, string> _, CancellationToken __)
        {
            call++;
            return Task.FromResult(batch);
        }

        var rows = await ConnectSecurePagedQuery.FetchAllPagesByIndexAsync(
            Fetch,
            new Dictionary<string, string>(),
            CancellationToken.None,
            pageSize: 100,
            maxPages: 10);

        Assert.Equal(5, rows.Count);
        Assert.Equal(1, call);
    }

    [Fact]
    public void SplitVersionTokens_DedupesSpaceSeparatedDuplicates()
    {
        var versions = PatchCatalogHelper.SplitVersionTokens("8.0.5 8.0.5 8.0.5 8.0.26.26169");
        Assert.Equal(2, versions.Count);
        Assert.Contains("8.0.5", versions);
        Assert.Contains("8.0.26.26169", versions);
    }

    [Fact]
    public void MeetsSeverityFilter_UsesMinimumRank()
    {
        Assert.True(PatchCatalogHelper.MeetsSeverityFilter("Critical", "high+"));
        Assert.True(PatchCatalogHelper.MeetsSeverityFilter("High", "high+"));
        Assert.False(PatchCatalogHelper.MeetsSeverityFilter("Medium", "high+"));
        Assert.False(PatchCatalogHelper.MeetsSeverityFilter("High", "critical"));
    }
}

public sealed class ConnectSecureCacheServiceTests
{
    [Fact]
    public void RemediationDatasetCache_ExpiresAfterInvalidate()
    {
        var cache = new ConnectSecureCacheService();
        var dataset = new RemediationDataset(42, DateTimeOffset.UtcNow, []);
        cache.SetRemediationDataset(42, dataset);

        Assert.True(cache.TryGetRemediationDataset(42, out var cached));
        Assert.Same(dataset, cached);

        cache.InvalidateRemediationDataset(42);
        Assert.False(cache.TryGetRemediationDataset(42, out _));
    }

    [Fact]
    public void PatchHostsCache_IsScopedBySolution()
    {
        var cache = new ConnectSecureCacheService();
        var hosts = new List<PatchAssetDetail>
        {
            new(1, "10.0.0.1", "host-a", 10, true, [], [], [])
        };

        cache.SetPatchHosts(7, 15, hosts);
        Assert.True(cache.TryGetPatchHosts(7, 15, out var cached));
        Assert.Single(cached);
        Assert.False(cache.TryGetPatchHosts(7, 16, out _));
    }

    [Fact]
    public void PatchProductHostsCache_IsScopedByProduct()
    {
        var cache = new ConnectSecureCacheService();
        var hosts = new List<PatchAssetDetail>
        {
            new(1, "10.0.0.1", "host-a", 10, true, [], [], [])
        };

        cache.SetPatchProductHosts(7, "MongoDB", hosts);
        Assert.True(cache.TryGetPatchProductHosts(7, "MongoDB", out var cached));
        Assert.Single(cached);
        Assert.False(cache.TryGetPatchProductHosts(7, "Firefox", out _));
    }

    [Fact]
    public void InvalidatePatchHosts_ClearsProductScopedCaches()
    {
        var cache = new ConnectSecureCacheService();
        cache.SetPatchProductHosts(42, "MongoDB", [new(1, "10.0.0.1", "host-a", 10, true, [], [], [])]);
        cache.SetProductRemediationAssetDetails(42, "MongoDB", [new(1, "10.0.0.1", "host-a", 10, true, ["MongoDB"], ["8.0.5"], [])]);

        cache.InvalidatePatchHosts(42);

        Assert.False(cache.TryGetPatchProductHosts(42, "MongoDB", out _));
        Assert.False(cache.TryGetProductRemediationAssetDetails(42, "MongoDB", out _));
    }

    [Fact]
    public void InvalidateCompany_ClearsPatchHosts()
    {
        var cache = new ConnectSecureCacheService();
        var hosts = new List<PatchAssetDetail>
        {
            new(1, "10.0.0.1", "host-a", 10, true, [], [], [])
        };

        cache.SetPatchHosts(42, 99, hosts);
        cache.SetRemediationDataset(42, new RemediationDataset(42, DateTimeOffset.UtcNow, []));
        Assert.True(cache.TryGetPatchHosts(42, 99, out _));

        cache.InvalidateCompany(42);

        Assert.False(cache.TryGetPatchHosts(42, 99, out _));
        Assert.False(cache.TryGetRemediationDataset(42, out _));
    }
}
