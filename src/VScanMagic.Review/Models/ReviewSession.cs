namespace VScanMagic.Review.Models;

public enum FindingStatus
{
    Open,
    InProgress,
    Remediated,
    Deferred,
    ClientAction,
    WontFix
}

public sealed class ReviewTask
{
    public string Text { get; set; } = "";
    public string? Owner { get; set; }
    public DateOnly? DueDate { get; set; }
}

public sealed class ReviewFinding
{
    public int Rank { get; set; }
    public int OriginalRank { get; set; }
    public string Product { get; set; } = "";
    public string Source { get; set; } = "Application";
    public double RiskScore { get; set; }
    public double Epss { get; set; }
    public double AvgCvss { get; set; }
    public int VulnCount { get; set; }
    public int Critical { get; set; }
    public int High { get; set; }
    public int Medium { get; set; }
    public int Low { get; set; }
    public List<ReviewAffectedSystem> AffectedSystems { get; set; } = [];
    public string CveIds { get; set; } = "";
    public string NvdEnrichment { get; set; } = "";
    public string OriginalFix { get; set; } = "";
    public string OriginalRemediation { get; set; } = "";
    public FindingStatus Status { get; set; } = FindingStatus.Open;
    public string RevisedRemediation { get; set; } = "";
    public string MeetingNotes { get; set; } = "";
    public List<ReviewTask> Tasks { get; set; } = [];
    public bool IncludeInExport { get; set; } = true;
    public bool ExcludedFromExport { get; set; }
    /// <summary>Technician explicitly added this finding from Network/Registry/reserve pool.</summary>
    public bool ManuallyPromoted { get; set; }
    /// <summary>ConnectSecure problem id resolved for CVE/registry/network suppress.</summary>
    public int? ConnectSecureProblemId { get; set; }
    /// <summary>ConnectSecure solution id used for application suppress.</summary>
    public int? ConnectSecureSolutionId { get; set; }
    /// <summary>ConnectSecure suppress_vulnerability record id for unsuppress.</summary>
    public int? ConnectSecureSuppressRecordId { get; set; }
    public bool ConnectSecureSuppressed { get; set; }
    public string? SuppressionReason { get; set; }
    public string? SuppressionComments { get; set; }
    public DateTimeOffset? SuppressedAt { get; set; }
    public decimal TimeEstimateHours { get; set; }
    public bool AfterHours { get; set; }
    public bool ThirdParty { get; set; }
    public bool TicketGenerated { get; set; }
    /// <summary>Tracks whether ThirdParty default was applied for legacy sessions.</summary>
    public bool TimeEstimateInitialized { get; set; }
}

public sealed class ReviewAffectedSystem
{
    public string HostName { get; set; } = "";
    public string Ip { get; set; } = "";
    public string Username { get; set; } = "";
    public int VulnCount { get; set; }
    public bool ExcludedFromExport { get; set; }
}

public sealed class ReviewSession
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string ClientName { get; set; } = "";
    public string ScanDate { get; set; } = "";
    public string? CompanyId { get; set; }
    public string Presenter { get; set; } = "";
    public bool IsRmitPlus { get; set; }
    public string? SourceFilePath { get; set; }
    public int ExportTopN { get; set; } = 10;
    /// <summary>SharePoint/OneDrive link to the Top N report for this quarterly deliverable.</summary>
    public string TopNReportUrl { get; set; } = "";
    /// <summary>SharePoint/OneDrive link to the quarterly reports folder.</summary>
    public string ReportsFolderUrl { get; set; } = "";
    /// <summary>TimeZest or other scheduling link for this deliverable (overrides Settings default).</summary>
    public string SchedulingLinkUrl { get; set; } = "";
    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.UtcNow;
    public List<ReviewFinding> Findings { get; set; } = [];
}
