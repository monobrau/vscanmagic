namespace VScanMagic.ConnectSecure;

/// <summary>
/// One row from <c>/r/report_queries/get_remediation</c>. Represents a single
/// (host × remediation) pair: which host needs which fix, with the metadata
/// needed to render and submit a patch.
/// </summary>
public sealed record RemediationRecord(
    int SolutionId,
    string Product,
    string Fix,
    string FixUrl,
    string Severity,
    string RemediationAction,
    string OsType,
    int AssetId,
    int AgentId,
    string HostName,
    string Ip,
    bool OnlineStatus,
    int TotalVulsCount,
    int CriticalVulsCount,
    int HighVulsCount,
    int MediumVulsCount,
    int LowVulsCount,
    double EpssVuls,
    string InstallSource,
    DateTimeOffset? FirstVulDiscovered,
    DateTimeOffset? LastVulDiscovered)
{
    public bool IsSoftwarePatch =>
        string.Equals(RemediationAction, "Software Patch", StringComparison.OrdinalIgnoreCase);

    public bool IsOsUpdate =>
        string.Equals(RemediationAction, "OS Update", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(RemediationAction, "OS Patch", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// All remediation rows for a company, fetched as one server-side filtered
/// page (<c>condition=company_id=X</c>). The dataset is the single source of
/// truth that powers Application Patching, OS Patching and the suppressible
/// remediation list.
/// </summary>
public sealed record RemediationDataset(
    int CompanyId,
    DateTimeOffset FetchedAt,
    IReadOnlyList<RemediationRecord> Records);
