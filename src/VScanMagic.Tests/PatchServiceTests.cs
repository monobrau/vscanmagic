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
}
