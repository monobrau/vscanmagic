using System.Text.RegularExpressions;
using VScanMagic.Core.Models;

namespace VScanMagic.Core.Risk;

public enum OsHostThresholdCategory
{
    Application,
    Windows11,
    WindowsOther,
    OtherOs
}

public static partial class OsHostVulnThresholdHelper
{
    private static readonly string[] NonOsProductMarkers =
    [
        "Visual C++", "Visual C#", ".NET", "Microsoft Edge", "Microsoft Office", "Microsoft Teams",
        "Google Chrome", "Mozilla Firefox", "Adobe", "Java ", "Java Runtime", "Installer",
        "Management Instrumentation", "Defender", "Silverlight", "SharePoint", "SQL Server",
        "Exchange", "OneDrive", "Skype", "Zoom", "Citrix", "FortiClient", "SonicWall",
        "Wireshark", "7-Zip", "WinRAR", "Notepad++", "PuTTY", "TeamViewer"
    ];

    /// <summary>
    /// OS host thresholds apply only to operating-system findings from the Application scan sheet —
    /// not Registry findings, Network findings, or third-party application products.
    /// </summary>
    public static bool IsOperatingSystemFinding(string? productName, string? source)
    {
        if (!IsApplicationScanSource(source))
            return false;

        return ClassifyProduct(productName) != OsHostThresholdCategory.Application;
    }

    public static bool IsApplicationScanSource(string? source)
    {
        var normalized = string.IsNullOrWhiteSpace(source) ? "Application" : source.Trim();
        return normalized.Equals("Application", StringComparison.OrdinalIgnoreCase);
    }

    public static OsHostThresholdCategory ClassifyProduct(string? productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return OsHostThresholdCategory.Application;

        var p = productName.Trim();
        if (LooksLikeApplicationProduct(p))
            return OsHostThresholdCategory.Application;

        if (IsWindows11Product(p))
            return OsHostThresholdCategory.Windows11;

        if (IsOtherWindowsOsProduct(p))
            return OsHostThresholdCategory.WindowsOther;

        if (IsOtherOsProduct(p))
            return OsHostThresholdCategory.OtherOs;

        return OsHostThresholdCategory.Application;
    }

    public static int GetThreshold(UserSettings settings, OsHostThresholdCategory category) =>
        category switch
        {
            OsHostThresholdCategory.Windows11 => settings.HostVulnThresholdWindows11,
            OsHostThresholdCategory.WindowsOther => settings.HostVulnThresholdWindowsOther,
            OsHostThresholdCategory.OtherOs => settings.HostVulnThresholdOtherOs,
            _ => 0
        };

    public static bool ShouldIncludeHost(string? productName, string? source, int hostVulnCount, UserSettings settings)
    {
        if (!IsOperatingSystemFinding(productName, source))
            return true;

        var category = ClassifyProduct(productName);
        var threshold = GetThreshold(settings, category);
        if (threshold <= 0)
            return true;

        return hostVulnCount >= threshold;
    }

    public static bool HasActiveThreshold(UserSettings settings, string? productName, string? source)
    {
        if (!IsOperatingSystemFinding(productName, source))
            return false;

        var category = ClassifyProduct(productName);
        return GetThreshold(settings, category) > 0;
    }

    private static bool LooksLikeApplicationProduct(string product) =>
        NonOsProductMarkers.Any(marker => product.Contains(marker, StringComparison.OrdinalIgnoreCase));

    private static bool IsWindows11Product(string product) =>
        product.Contains("Windows 11", StringComparison.OrdinalIgnoreCase) ||
        Windows11ProductPattern().IsMatch(product);

    private static bool IsOtherWindowsOsProduct(string product)
    {
        if (product.Contains("Windows 11", StringComparison.OrdinalIgnoreCase))
            return false;

        return WindowsOsProductPattern().IsMatch(product) ||
               product.Contains("Microsoft Windows", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOtherOsProduct(string product)
    {
        string[] otherOsMarkers =
        [
            "Ubuntu", "Red Hat", "RHEL", "macOS", "Mac OS", "Debian", "CentOS",
            "SUSE", "Unix", "AIX", "Solaris", "FreeBSD", "OpenBSD", "VMware ESXi", "ESXi"
        ];

        return otherOsMarkers.Any(marker => product.Contains(marker, StringComparison.OrdinalIgnoreCase)) &&
               !LooksLikeApplicationProduct(product);
    }

    [GeneratedRegex(@"^Windows\s+11", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex Windows11ProductPattern();

    [GeneratedRegex(@"^Windows\s+(10|Server|8|7|\d{4})", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex WindowsOsProductPattern();
}
