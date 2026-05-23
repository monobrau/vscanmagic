using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class ExcelCellText
{
    public const int MaxLength = 32767;

    public static string Truncate(string? text, string suffix = "…")
    {
        if (string.IsNullOrEmpty(text))
            return "";

        if (text.Length <= MaxLength)
            return text;

        var keep = Math.Max(0, MaxLength - suffix.Length);
        return text[..keep] + suffix;
    }

    public static string FormatAffectedSystemsForExcel(ReviewFinding finding)
    {
        var systems = FindingExportDetails.IncludedSystems(finding);
        if (systems.Count == 0)
            return "(none included after review exclusions)";

        var inline = FindingExportDetails.FormatAffectedSystemsInline(finding);
        if (inline.Length <= MaxLength - 64)
            return inline;

        var parts = new List<string>(systems.Count);
        foreach (var system in systems)
        {
            var part = FindingExportDetails.FormatAffectedSystem(system);
            var candidate = parts.Count == 0 ? part : string.Join(", ", parts) + ", " + part;
            if (candidate.Length > MaxLength - 80)
                break;
            parts.Add(part);
        }

        var result = string.Join(", ", parts);
        if (parts.Count < systems.Count)
            result += $" … ({systems.Count} hosts total — see Affected Systems sheet)";

        return Truncate(result);
    }
}
