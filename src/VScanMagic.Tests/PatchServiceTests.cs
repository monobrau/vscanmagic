using System.Text.Json;
using VScanMagic.ConnectSecure;

namespace VScanMagic.Tests;

public sealed class ConnectSecurePatchServiceTests
{
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
    public void RemediationPlanCache_ExpiresAfterInvalidate()
    {
        var cache = new ConnectSecureCacheService();
        var rows = new List<JsonElement>();
        cache.SetRemediationPlan(42, rows);

        Assert.True(cache.TryGetRemediationPlan(42, out var cached));
        Assert.Same(rows, cached);

        cache.InvalidateRemediationPlan(42);
        Assert.False(cache.TryGetRemediationPlan(42, out _));
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
}
