using System.Text.Json;
using VScanMagic.ConnectSecure;

namespace VScanMagic.Tests;

public sealed class RemediationDatasetParsingTests
{
    private static JsonElement ParseElement(string json) =>
        JsonDocument.Parse(json).RootElement.Clone();

    [Fact]
    public void ParseRecord_PopulatesAllFields()
    {
        var row = ParseElement(@"{
            ""solution_id"": 41,
            ""os_type"": ""windows"",
            ""fix"": ""4.6.6"",
            ""url"": ""https://www.wireshark.org/download.html"",
            ""product"": ""Wireshark"",
            ""remediation_action"": ""Software Patch"",
            ""severity"": ""Low"",
            ""epss_vuls"": 0.0042,
            ""critical_vuls_count"": 0,
            ""high_vuls_count"": 0,
            ""medium_vuls_count"": 0,
            ""low_vuls_count"": 1,
            ""total_vuls_count"": 1,
            ""first_vul_discovered"": ""2026-05-22T08:02:45.719371"",
            ""last_vul_discovered"": ""2026-05-22T08:02:45.719371"",
            ""install_source"": """",
            ""agent_id"": 402706,
            ""asset_id"": 24076207,
            ""company_id"": 17624,
            ""company_name"": ""Stream Jog"",
            ""host_name"": ""DC-05.dorks.lan"",
            ""ip_addresses"": ""192.168.54.3"",
            ""online_status"": true
        }");

        var record = ConnectSecureRemediationService.ParseRecord(row);

        Assert.Equal(41, record.SolutionId);
        Assert.Equal("Wireshark", record.Product);
        Assert.Equal("4.6.6", record.Fix);
        Assert.Equal("Low", record.Severity);
        Assert.Equal("Software Patch", record.RemediationAction);
        Assert.Equal("windows", record.OsType);
        Assert.Equal(24076207, record.AssetId);
        Assert.Equal(402706, record.AgentId);
        Assert.Equal("DC-05.dorks.lan", record.HostName);
        Assert.Equal("192.168.54.3", record.Ip);
        Assert.True(record.OnlineStatus);
        Assert.True(record.IsSoftwarePatch);
        Assert.False(record.IsOsUpdate);
        Assert.Equal(1, record.LowVulsCount);
        Assert.Equal(1, record.TotalVulsCount);
        Assert.Equal(0.0042, record.EpssVuls, 4);
    }

    [Fact]
    public void ParseRecord_RecognizesOsUpdate()
    {
        var row = ParseElement(@"{
            ""solution_id"": 999,
            ""product"": ""Windows 1125H2"",
            ""fix"": ""5089549"",
            ""remediation_action"": ""OS Update"",
            ""severity"": ""Critical"",
            ""os_type"": ""windows""
        }");

        var record = ConnectSecureRemediationService.ParseRecord(row);

        Assert.True(record.IsOsUpdate);
        Assert.False(record.IsSoftwarePatch);
        Assert.Equal("Windows 1125H2", record.Product);
        Assert.Equal("5089549", record.Fix);
    }

    [Fact]
    public void ParseRecord_FallsBackToIdWhenAssetIdMissing()
    {
        var row = ParseElement(@"{ ""id"": 12345, ""product"": ""Test"" }");
        var record = ConnectSecureRemediationService.ParseRecord(row);
        Assert.Equal(12345, record.AssetId);
    }

    [Fact]
    public void ParseRecord_PicksFirstIpFromCsvList()
    {
        var row = ParseElement(@"{ ""ip_addresses"": ""10.0.0.5, 10.0.0.6"", ""product"": ""x"" }");
        var record = ConnectSecureRemediationService.ParseRecord(row);
        Assert.Equal("10.0.0.5", record.Ip);
    }

    [Fact]
    public void ParseRecord_HandlesStringNumbers()
    {
        var row = ParseElement(@"{ ""solution_id"": ""41"", ""asset_id"": ""24076207"", ""epss_vuls"": ""0.5"" }");
        var record = ConnectSecureRemediationService.ParseRecord(row);

        Assert.Equal(41, record.SolutionId);
        Assert.Equal(24076207, record.AssetId);
        Assert.Equal(0.5, record.EpssVuls, 4);
    }
}
