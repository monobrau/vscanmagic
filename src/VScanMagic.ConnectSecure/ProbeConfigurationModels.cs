namespace VScanMagic.ConnectSecure;

public sealed class ProbeConfigurationData
{
    public int CompanyId { get; init; }
    public string CompanyName { get; init; } = "";
    public List<ProbeAgentConfiguration> Probes { get; } = [];
    public List<CompanyCredentialEntry> Credentials { get; } = [];
    public List<IntegrationCredentialEntry> Integrations { get; } = [];
    public List<CompanyIntegrationMappingEntry> IntegrationMappings { get; } = [];
    public List<InternalSubnetEntry> InternalSubnets { get; } = [];
    public IReadOnlyList<string> KnownCredentialTypes { get; set; } = [];
    public IReadOnlyList<string> KnownIntegrationNames { get; set; } = [];
}

public sealed record ProbeAgentConfiguration(
    int AgentId,
    string HostName,
    string Ip,
    IReadOnlyList<int> CredentialIds,
    IReadOnlyList<int> DiscoverySettingIds);

public sealed record CompanyCredentialEntry(
    int Id,
    string Name,
    string CredentialType,
    string OsType,
    string AddressType,
    string Address,
    bool IsValid,
    string FailureReason,
    string ParamsSummary,
    IReadOnlyList<int> MappedAgentIds);

public sealed record IntegrationCredentialEntry(
    int Id,
    string Name,
    string IntegrationName,
    string TicketUrl,
    string ParamsSummary);

public sealed record CompanyIntegrationMappingEntry(
    int Id,
    string IntegrationName,
    string SourceCompanyName,
    string DestCompanyName,
    string SiteName,
    int? CredentialId);

public sealed record CredentialSaveRequest(
    string Name,
    string CredentialType,
    string? OsType,
    string? AddressType,
    string? Address,
    string ParamsJson,
    bool MergeExistingSecrets);

public sealed record IntegrationCredentialSaveRequest(
    string Name,
    string IntegrationName,
    string? TicketUrl,
    string ParamsJson,
    bool MergeExistingSecrets);

public sealed record IntegrationMappingSaveRequest(
    string IntegrationName,
    string? SourceCompanyName,
    string? DestCompanyName,
    string? DestCompanyId,
    string? SiteName,
    string? SiteId,
    int? CredentialId,
    string ParamsJson,
    bool MergeExistingFields);
