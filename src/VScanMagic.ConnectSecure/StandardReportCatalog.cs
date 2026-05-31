namespace VScanMagic.ConnectSecure;

public static class StandardReportCatalog
{
    public static readonly IReadOnlyList<StandardReportRequest> AllVulnerabilitiesOnly =
    [
        new("all-vulnerabilities", "All Vulnerabilities Report", "xlsx")
    ];

    /// <summary>Standard bundle minus All Vulnerabilities — fetched at export time, not during review prep.</summary>
    public static readonly IReadOnlyList<StandardReportRequest> SupplementalCompanyReports =
    [
        new("suppressed-vulnerabilities", "Suppressed Vulnerabilities", "xlsx"),
        new("external-vulnerabilities", "External Scan", "xlsx"),
        new("network-vulnerabilities", "Network Vulnerabilities", "xlsx"),
        new("executive-summary", "Executive Summary Report", "docx"),
        new("pending-epss", "Pending Remediation EPSS Score Reports", "xlsx")
    ];

    public static readonly IReadOnlyList<StandardReportRequest> DefaultCompanyReports =
    [
        new("all-vulnerabilities", "All Vulnerabilities Report", "xlsx"),
        new("suppressed-vulnerabilities", "Suppressed Vulnerabilities", "xlsx"),
        new("external-vulnerabilities", "External Scan", "xlsx"),
        new("network-vulnerabilities", "Network Vulnerabilities", "xlsx"),
        new("executive-summary", "Executive Summary Report", "docx"),
        new("pending-epss", "Pending Remediation EPSS Score Reports", "xlsx")
    ];

    public static readonly Dictionary<string, string> CategoryPatterns = new(StringComparer.OrdinalIgnoreCase)
    {
        ["all-vulnerabilities"] = "all vulnerabilities report",
        ["suppressed-vulnerabilities"] = "suppressed vulnerabilities",
        ["external-vulnerabilities"] = "external scan",
        ["executive-summary"] = "executive summary report",
        ["pending-epss"] = "pending remediation epss score reports",
        ["network-vulnerabilities"] = "network scan findings"
    };

    public static readonly Dictionary<string, string> KnownReportIds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["all-vulnerabilities"] = "00000000000000000000000000000000",
        ["suppressed-vulnerabilities"] = "1d091564830b44c485a0ddc35ace9ac6",
        ["external-vulnerabilities"] = "01beb6b930744e11b690bb9dc25118fb",
        ["executive-summary"] = "1cd4f45884264d15bee4173dc58b6a57",
        ["pending-epss"] = "85d4913c0dbc4fc782b858f0d27dd180"
    };

    public static readonly Dictionary<string, string> DisplayNames = new(StringComparer.OrdinalIgnoreCase)
    {
        ["all-vulnerabilities"] = "All Vulnerabilities Report",
        ["suppressed-vulnerabilities"] = "Suppressed Vulnerabilities",
        ["external-vulnerabilities"] = "External Scan",
        ["executive-summary"] = "Executive Summary Report",
        ["pending-epss"] = "Pending Remediation EPSS Score Reports",
        ["network-vulnerabilities"] = "Network Scan Findings"
    };
}

public sealed record StandardReportRequest(string Type, string Name, string Extension);

public sealed record StandardReportDownloadResult(
    IReadOnlyList<DownloadedReport> Succeeded,
    IReadOnlyList<FailedReport> Failed);

public sealed record DownloadedReport(string Type, string Name, string Path);

public sealed record FailedReport(string Type, string Name, string Error);

public sealed class StandardReportDescriptor
{
    public required string Id { get; init; }
    public required string ReportType { get; init; }
    public string Category { get; init; } = "";
    public string CategoryDisplay { get; init; } = "";
    public string DisplayName { get; init; } = "";
}
