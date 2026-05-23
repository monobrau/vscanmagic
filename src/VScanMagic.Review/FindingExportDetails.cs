using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class FindingExportDetails
{
    public static string GetAffectedSystemIdentifier(ReviewAffectedSystem system)
    {
        if (!string.IsNullOrWhiteSpace(system.HostName))
            return system.HostName.Trim();
        if (!string.IsNullOrWhiteSpace(system.Ip))
            return system.Ip.Trim();
        return "";
    }

    public static string FormatAffectedSystemCompact(ReviewAffectedSystem system)
    {
        var line = GetAffectedSystemIdentifier(system);
        if (string.IsNullOrWhiteSpace(line))
            return "";

        if (!string.IsNullOrWhiteSpace(system.Username))
            line += $" ({system.Username.Trim()})";

        var ip = system.Ip.Trim();
        if (!string.IsNullOrWhiteSpace(ip) &&
            !line.Equals(ip, StringComparison.OrdinalIgnoreCase))
            line += $" - {ip}";

        return line;
    }

    public static string FormatAffectedSystem(ReviewAffectedSystem system)
    {
        var line = FormatAffectedSystemCompact(system);
        if (string.IsNullOrWhiteSpace(line))
            return "";

        if (system.VulnCount > 0)
            line += $" [{system.VulnCount} vulns]";

        return line;
    }

    public static string FormatAffectedSystemsInline(ReviewFinding finding) =>
        string.Join(", ", IncludedSystems(finding).Select(FormatAffectedSystem));

    public static string FormatAffectedSystemsCompactInline(ReviewFinding finding) =>
        string.Join(", ", IncludedSystems(finding).Select(FormatAffectedSystemCompact).Where(x => !string.IsNullOrWhiteSpace(x)));

    public static string FormatAffectedSystemsMultiline(ReviewFinding finding) =>
        string.Join(Environment.NewLine,
            IncludedSystems(finding).Select(s => $"- {FormatAffectedSystem(s)}"));

    public static IReadOnlyList<ReviewAffectedSystem> IncludedSystems(ReviewFinding finding) =>
        finding.IncludedSystems()
            .GroupBy(s => $"{s.HostName}|{s.Ip}|{s.Username}", StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(x => x.VulnCount).First())
            .OrderByDescending(s => s.VulnCount)
            .ToList();

    public static IReadOnlyList<string> GetCveIds(ReviewFinding finding) =>
        CveReferenceHelper.SplitCveIds(finding.CveIds);

    public static string FormatCveIds(ReviewFinding finding) =>
        string.Join("; ", GetCveIds(finding));

    public static string FormatReferenceLinks(ReviewFinding finding, string? separator = null) =>
        CveReferenceHelper.FormatReferenceLinks(finding.CveIds, separator);

    public static string GetRemediationText(ReviewFinding finding) =>
        string.IsNullOrWhiteSpace(finding.RevisedRemediation)
            ? finding.OriginalRemediation
            : finding.RevisedRemediation;
}
