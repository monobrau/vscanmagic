namespace VScanMagic.ConnectSecure;

public enum ConnectSecurePatchType
{
    App,
    Os
}

public enum ConnectSecurePatchWhen
{
    Now,
    Later
}

public sealed record PatchableApplicationEntry(
    int SolutionId,
    string Product,
    string Fix,
    bool IsPatchable,
    int AffectedAssets,
    IReadOnlyList<int> AssetIds,
    string Severity,
    string RemediationAction);

public sealed record PatchAssetDetail(
    int AssetId,
    string Ip,
    string HostName,
    int AgentId,
    bool OnlineStatus,
    IReadOnlyList<string> ApplicationNames,
    IReadOnlyList<string> Versions,
    IReadOnlyList<string> Paths);

public class ApplicationPatchRequest
{
    public int CompanyId { get; set; }
    public IReadOnlyList<int> AssetIds { get; set; } = [];
    public IReadOnlyList<int> AgentIds { get; set; } = [];
    public IReadOnlyList<string> IncludedApplications { get; set; } = [];
    public IReadOnlyList<string> ExcludedApplications { get; set; } = [];
    public IReadOnlyList<string> IncludeTags { get; set; } = [];
    public IReadOnlyList<string> ExcludeTags { get; set; } = [];
    public IReadOnlyList<string> TargetHostNames { get; set; } = [];
    public Dictionary<string, string> FromVersions { get; set; } = new(StringComparer.Ordinal);
    public ConnectSecurePatchType PatchType { get; set; } = ConnectSecurePatchType.App;
    public bool TriggerReboot { get; set; }
}

public sealed class ScheduledApplicationPatchRequest : ApplicationPatchRequest
{
    public DateTime ScheduledAt { get; set; }
}

public sealed record PatchOperationResult(bool Success, string Message);

public enum HostPatchStatus
{
    Unknown,
    Offline,
    EndOfLife,
    AtTarget,
    Pending
}

public sealed record PatchHostView(
    PatchAssetDetail Detail,
    HostPatchStatus Status,
    string StatusLabel);

public sealed record PatchJobEntry(
    string JobId,
    string Type,
    string Status,
    string Description,
    string? HostName,
    string? AgentIp,
    DateTimeOffset? Updated);

public sealed record OsPendingPatchEntry(
    string OsName,
    string OsVersion,
    string Fix,
    int AffectedAssets,
    IReadOnlyList<int> AssetIds);

public sealed record SuppressibleRemediationEntry(
    int SolutionId,
    string Product,
    string Fix,
    string Severity,
    string RemediationAction,
    bool IsPatchable,
    int AffectedAssets);

public sealed record SuppressVulnerabilityRequest
{
    public int CompanyId { get; set; }
    public int SolutionId { get; set; }
    public int AssetId { get; set; }
    public string Product { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Comments { get; set; } = "";
}

public sealed record ScanTriggerResult(bool Success, string Message, int TriggeredCount);
