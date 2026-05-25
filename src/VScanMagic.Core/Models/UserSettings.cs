namespace VScanMagic.Core.Models;

public sealed class UserSettings
{
    public string PreparedBy { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string CompanyAddress { get; set; } = "";
    public string Email { get; set; } = "";
    public string PhoneNumber { get; set; } = "";
    public string CompanyPhoneNumber { get; set; } = "";
    public string? SettingsDirectory { get; set; }
    public double FilterMinEpss { get; set; }
    public bool FilterIncludeCritical { get; set; } = true;
    public bool FilterIncludeHigh { get; set; } = true;
    public bool FilterIncludeMedium { get; set; }
    public bool FilterIncludeLow { get; set; }
    public string FilterTopN { get; set; } = "10";
    public string LastOutputDirectory { get; set; } = "";
    public string ReportsBasePath { get; set; } = "";

    /// <summary>Hosts with fewer vulns than this are excluded for Windows 11 findings. 0 = off.</summary>
    [System.Text.Json.Serialization.JsonPropertyName("HostnameReviewWindows11Threshold")]
    public int HostVulnThresholdWindows11 { get; set; } = 350;

    /// <summary>Windows 10, Windows Server, etc. 0 = off.</summary>
    public int HostVulnThresholdWindowsOther { get; set; }

    /// <summary>Linux, macOS, and other non-Windows OS findings. 0 = off.</summary>
    public int HostVulnThresholdOtherOs { get; set; }

    /// <summary>Optional NVD API key for higher rate limits when enriching CVE details.</summary>
    public string NvdApiKey { get; set; } = "";

    /// <summary>Fetch NVD summaries for CVE findings when starting a review session.</summary>
    public bool NvdEnrichOnIngest { get; set; } = true;

    /// <summary>Default TimeZest or other scheduling link inserted into deliverable email templates.</summary>
    public string SchedulingLinkUrl { get; set; } = "";
}
