using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class FindingRemediationExport
{
    public static bool IsRemediationEdited(ReviewFinding finding)
    {
        var original = Normalize(finding.OriginalRemediation);
        if (original.Length == 0)
            return false;

        var revised = Normalize(finding.RevisedRemediation);
        return revised.Length > 0 && !string.Equals(revised, original, StringComparison.Ordinal);
    }

    public static string GetWordRemediationText(ReviewFinding finding, RemediationRuleService remediationRules) =>
        CveExportFormatter.GetRemediationInstructions(finding, remediationRules, forWord: true);

    public static string GetTicketRemediationText(ReviewFinding finding, RemediationRuleService remediationRules) =>
        CveExportFormatter.GetRemediationInstructions(finding, remediationRules, forWord: false);

    public static string GetTimeEstimateRemediationText(ReviewFinding finding, RemediationRuleService remediationRules)
    {
        if (IsRemediationEdited(finding))
            return CveEnrichmentPolicy.AppendNvdContext(FindingExportDetails.GetRemediationText(finding), finding);

        var guidance = remediationRules.GetGuidance(finding.Product, forWord: false);
        return CveEnrichmentPolicy.AppendNvdContext(guidance, finding);
    }

    public static string? GetConnectSecureSolution(ReviewFinding finding)
    {
        var fix = finding.OriginalFix?.Trim();
        if (string.IsNullOrWhiteSpace(fix))
            return null;

        var remediation = Normalize(FindingExportDetails.GetRemediationText(finding));
        var fixNorm = Normalize(fix);
        if (fixNorm.Length == 0)
            return null;

        if (string.Equals(fixNorm, remediation, StringComparison.Ordinal))
            return null;

        if (remediation.Contains(fixNorm, StringComparison.Ordinal))
            return null;

        return fix;
    }

    private static string Normalize(string? text) =>
        string.IsNullOrWhiteSpace(text)
            ? ""
            : string.Join(' ', text.Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
