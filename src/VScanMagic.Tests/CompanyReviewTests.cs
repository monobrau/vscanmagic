using System.Text.Json;
using VScanMagic.ConnectSecure;
using VScanMagic.Core.Services;

namespace VScanMagic.Tests;

public sealed class CompanyReviewServiceTests
{
    [Fact]
    public void CompanyQuery_UsesConditionNotCompanyIdParam()
    {
        var query = ConnectSecureCompanyReviewService.CompanyQuery(15998, limit: 500, orderBy: "updated desc");
        Assert.Equal("company_id=15998", query["condition"]);
        Assert.Equal("500", query["limit"]);
        Assert.Equal("0", query["skip"]);
        Assert.Equal("updated desc", query["order_by"]);
        Assert.False(query.ContainsKey("company_id"));
    }

    [Fact]
    public void LightweightAssetsQuery_UsesCompanyIdParam()
    {
        var query = ConnectSecureCompanyReviewService.LightweightAssetsQuery(15998);
        Assert.Equal("15998", query["company_id"]);
        Assert.Equal("5000", query["limit"]);
    }

    [Fact]
    public void ProbeAgentsQuery_MatchesPowerShellFilter()
    {
        var query = ConnectSecureCompanyReviewService.ProbeAgentsQuery(42);
        Assert.Contains("agent_type='PROBE'", query["condition"]);
        Assert.Contains("company_id=42", query["condition"]);
    }

    [Fact]
    public void BuildChecks_MarksOffline30AsWarningWhenPresent()
    {
        var data = new CompanyReviewData { AgentsOffline30PlusDays = 2 };
        var checks = ConnectSecureCompanyReviewService.BuildChecks(data);
        var offline = checks.First(c => c.Label.StartsWith("4."));
        Assert.False(offline.Ok);
    }

    [Fact]
    public void CombineSubnetLines_MergesProbeSubnetsAndScanTargets()
    {
        var data = new CompanyReviewData();
        data.ProbesSubnets.Add("10.0.0.0/24");
        data.ScanTargets.Add("203.0.113.5");
        data.ScanTargets.Add("10.0.0.0/24");

        var lines = ConnectSecureCompanyReviewService.CombineSubnetLines(data);
        Assert.Equal(2, lines.Count);
        Assert.Contains("10.0.0.0/24", lines);
        Assert.Contains("203.0.113.5", lines);
    }

    [Fact]
    public void RebuildQuickWins_RemovesNmapRecommendationAfterConfigured()
    {
        var data = new CompanyReviewData { ProbesWithBoth = 1 };
        data.ProbeAgentsNmapInfo.Add(new ProbeNmapInfo(1, "probe1", "10.0.0.1", "(not set)", null, []));
        ConnectSecureCompanyReviewService.RebuildQuickWins(data);
        Assert.Contains(data.QuickWins, w => w.Contains("nmap interface", StringComparison.OrdinalIgnoreCase));

        data.ProbeAgentsNmapInfo[0] = data.ProbeAgentsNmapInfo[0] with { NmapInterface = "eth0 (10.0.0.1)" };
        ConnectSecureCompanyReviewService.RebuildQuickWins(data);
        Assert.DoesNotContain(data.QuickWins, w => w.Contains("nmap interface", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void IsExternalDiscoverySetting_ReturnsTrueForExternalScan()
    {
        using var doc = JsonDocument.Parse("""{"discovery_settings_type":"externalscan"}""");
        Assert.True(ConnectSecureCompanyReviewService.IsExternalDiscoverySetting(doc.RootElement));
    }

    [Fact]
    public void IsExternalDiscoverySetting_ReturnsFalseForInternalNetwork()
    {
        using var doc = JsonDocument.Parse("""{"discovery_settings_type":"networkscan"}""");
        Assert.False(ConnectSecureCompanyReviewService.IsExternalDiscoverySetting(doc.RootElement));
    }
}

public sealed class ProbeInterfaceHelperTests
{
    [Fact]
    public void ParseAvailableInterfaces_ReadsStringArray()
    {
        using var doc = JsonDocument.Parse("""{"interfaces":["eth0","eth1"]}""");
        var parsed = ProbeInterfaceHelper.ParseAvailableInterfaces(doc.RootElement);
        Assert.Equal(["eth0", "eth1"], parsed);
    }

    [Fact]
    public void ParseAvailableInterfaces_ReadsObjectArrayWithNameAndIp()
    {
        using var doc = JsonDocument.Parse("""
            {"interfaces":[{"name":"Ethernet","ip":"10.0.0.5"},{"device":"Wi-Fi","address":"192.168.1.10"}]}
            """);
        var parsed = ProbeInterfaceHelper.ParseAvailableInterfaces(doc.RootElement);
        Assert.Equal(["Ethernet (10.0.0.5)", "Wi-Fi (192.168.1.10)"], parsed);
    }

    [Fact]
    public void BuildDropdownOptions_IncludesCurrentValueWhenMissingFromList()
    {
        var options = ProbeInterfaceHelper.BuildDropdownOptions(["eth0 (10.0.0.1)"], "eth1 (10.0.0.2)");
        Assert.Equal(2, options.Count);
        Assert.Contains(options, o => o.Value == "eth1 (10.0.0.2)");
    }

    [Fact]
    public void BuildDropdownOptions_ReturnsNotSetWhenEmpty()
    {
        var options = ProbeInterfaceHelper.BuildDropdownOptions([], null);
        Assert.Single(options);
        Assert.Equal("", options[0].Value);
        Assert.Equal("(not set)", options[0].Label);
    }
}

public sealed class ExternalSubnetHelperTests
{
    [Fact]
    public void ValidateExternalTargets_FlagsNetworkGatewayAndBroadcast()
    {
        var issues = ExternalSubnetHelper.ValidateExternalTargets(
            ["192.168.1.0", "192.168.1.1", "192.168.1.255"],
            "192.168.1.0/24");

        Assert.Equal(3, issues.Count);
        Assert.Contains(issues, i => i.Contains("network", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(issues, i => i.Contains("gateway", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(issues, i => i.Contains("broadcast", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ExpandCidrToUsableIps_SkipsNetworkGatewayAndBroadcastForSlash24()
    {
        var ips = ExternalSubnetHelper.ExpandCidrToUsableIps("192.168.1.0/24");
        Assert.Equal(253, ips.Count);
        Assert.Equal("192.168.1.2", ips[0]);
        Assert.Equal("192.168.1.254", ips[^1]);
        Assert.DoesNotContain("192.168.1.0", ips);
        Assert.DoesNotContain("192.168.1.1", ips);
        Assert.DoesNotContain("192.168.1.255", ips);
    }

    [Fact]
    public void ExpandCidrToUsableIps_Slash29HasFiveUsableHosts()
    {
        var ips = ExternalSubnetHelper.ExpandCidrToUsableIps("203.0.113.0/29");
        Assert.Equal(["203.0.113.2", "203.0.113.3", "203.0.113.4", "203.0.113.5", "203.0.113.6"], ips);
    }

    [Fact]
    public void ParseAndValidateScanInput_RejectsExplicitGatewayInCidr()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.1.0/24, 192.168.1.1");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("gateway", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ParseAndValidateScanInput_ExpandsCidrToTargetIpList()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("203.0.113.0/29");
        Assert.True(result.IsValid);
        Assert.Equal("203.0.113.0/29", result.Address);
        Assert.Equal(5, result.ScanIps.Count);
        Assert.StartsWith("203.0.113.2", result.TargetIp);
    }

    [Fact]
    public void ParseAndValidateScanInput_AllowsSinglePublicIp()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("203.0.113.50");
        Assert.True(result.IsValid);
        Assert.Equal(["203.0.113.50"], result.ScanIps);
    }

    [Fact]
    public void ParseAndValidateScanInput_AcceptsDottedMaskWithPipeSeparator()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.54.0 | 255.255.255.0");
        Assert.True(result.IsValid);
        Assert.Equal("192.168.54.0/24", result.Address);
        Assert.Equal(253, result.ScanIps.Count);
    }

    [Fact]
    public void ParseAndValidateScanInput_AcceptsDottedMaskWithSlash()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.54.0/255.255.255.0");
        Assert.True(result.IsValid);
        Assert.Equal("192.168.54.0/24", result.Address);
    }

    [Fact]
    public void ParseAndValidateScanInput_AcceptsSpaceSeparatedMask()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.54.0 255.255.255.0");
        Assert.True(result.IsValid);
        Assert.Equal("192.168.54.0/24", result.Address);
    }

    [Fact]
    public void ParseAndValidateScanInput_RejectsIncompleteDottedMask()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.54.50 | 255.255.255");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("four octets", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ParseAndValidateScanInput_RejectsIncompleteCidrPrefix()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.54.50/");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("prefix length", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ParseAndValidateScanInput_AcceptsVlan50Cidr()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("192.168.50.0/24");
        Assert.True(result.IsValid);
        Assert.Equal("192.168.50.0/24", result.Address);
        Assert.Equal(253, result.ScanIps.Count);
    }

    [Fact]
    public void DescribeExpandedRange_AcceptsVlan50Cidr()
    {
        var description = ExternalSubnetHelper.DescribeExpandedRange("192.168.50.0/24");
        Assert.Contains("253", description);
    }

    [Fact]
    public void ValidateScanInputForUi_RejectsPartialSubnetNotation()
    {
        var errors = ExternalSubnetHelper.ValidateScanInputForUi("192.168.50.0", strict: true);
        Assert.Empty(errors);

        errors = ExternalSubnetHelper.ValidateScanInputForUi("192.168.50.0 255.255.255.0", strict: true);
        Assert.Empty(errors);

        errors = ExternalSubnetHelper.ValidateScanInputForUi("192.168.50.0/24", strict: true);
        Assert.Empty(errors);
    }

    [Fact]
    public void ValidateScanInputForUi_FlagsSubnetLikeInputThatOnlyResolvesSingleIp()
    {
        var errors = ExternalSubnetHelper.ValidateScanInputForUi("192.168.50.0 255.255.255", strict: true);
        Assert.NotEmpty(errors);
        Assert.Contains(errors, e =>
            e.Contains("Could not parse a subnet", StringComparison.OrdinalIgnoreCase) ||
            e.Contains("subnet mask", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ParseAndValidateScanInput_RejectsStandaloneSubnetMask()
    {
        var result = ExternalSubnetHelper.ParseAndValidateScanInput("255.255.255.0");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("subnet mask", StringComparison.OrdinalIgnoreCase));
    }
}

public sealed class CredentialTypeCatalogTests
{
    [Fact]
    public void MergeKnownTypes_IncludesDefaultsAndDiscovered()
    {
        var types = CredentialTypeCatalog.MergeKnownTypes(["customldap", "windows"]);
        Assert.Contains("windows", types);
        Assert.Contains("customldap", types);
        Assert.Contains("snmp", types);
    }

    [Fact]
    public void Resolve_UnknownTypeUsesGenericFields()
    {
        var def = CredentialTypeCatalog.Resolve("customldap");
        Assert.Equal("customldap", def.Type);
        Assert.Empty(def.Fields);
    }
}

public sealed class ConnectSecureParamsHelperTests
{
    [Fact]
    public void MergeParamsJson_PreservesExistingSecretWhenIncomingBlank()
    {
        var existing = ConnectSecureParamsHelper.ParseParamsObject("""{"password":"secret","username":"admin"}""");
        var incoming = ConnectSecureParamsHelper.ParseParamsObject("""{"password":"","username":"admin2"}""");
        var merged = ConnectSecureParamsHelper.MergeParamsJson(existing, incoming, mergeExistingSecrets: true);
        Assert.Equal("secret", merged["password"]!.GetValue<string>());
        Assert.Equal("admin2", merged["username"]!.GetValue<string>());
    }
}

public sealed class CoveredSoftwareServiceTests
{
    [Fact]
    public void GetDefaults_IncludesMicrosoftPattern()
    {
        var defaults = CoveredSoftwareService.GetDefaults();
        Assert.Contains(defaults, d => d.Pattern.Contains("Microsoft", StringComparison.Ordinal));
    }
}
