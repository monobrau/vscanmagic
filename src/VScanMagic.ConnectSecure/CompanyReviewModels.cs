namespace VScanMagic.ConnectSecure;

public sealed class CompanyReviewData
{
    public int CompanyId { get; init; }
    public string CompanyName { get; init; } = "";
    public int AgentCount { get; set; }
    public int ProbesWithCredentials { get; set; }
    public int ProbesWithNetworks { get; set; }
    public int ProbesWithBoth { get; set; }
    public List<string> ProbesSubnets { get; } = [];
    public List<string> ScanTargets { get; } = [];
    public List<ExternalAssetEntry> ExternalAssets { get; } = [];
    public List<InternalSubnetEntry> InternalSubnets { get; } = [];
    public List<string> SubnetIssues { get; } = [];
    public int AgentsOffline7PlusDays { get; set; }
    public int AgentsOffline14PlusDays { get; set; }
    public int AgentsOffline30PlusDays { get; set; }
    public List<string> AgentsOffline30PlusNames { get; } = [];
    public bool FirewallActive { get; set; }
    public int FirewallCount { get; set; }
    public string FirewallType { get; set; } = "";
    public string? LastInternalScan { get; set; }
    public string? LastExternalScan { get; set; }
    public List<ProbeNmapInfo> ProbeAgentsNmapInfo { get; } = [];
    public List<string> QuickWins { get; } = [];
}

public sealed record InternalSubnetEntry(
    int? DiscoverySettingId,
    int? MappingId,
    int? ProbeAgentId,
    string ProbeHostName,
    string Name,
    string Address,
    string? TargetIp,
    int ScanIpCount);

public sealed record ExternalAssetEntry(
    int? Id,
    string Name,
    string Address,
    string? TargetIp,
    int ScanIpCount);

public sealed record ProbeNmapInfo(
    int? AgentId,
    string HostName,
    string Ip,
    string NmapInterface,
    string? Port,
    IReadOnlyList<string> AvailableInterfaces);

public sealed record CompanyReviewCheck(string Label, string Value, bool Ok);
