using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class CveEnrichmentPolicy
{
    public static IReadOnlyList<string> GetCveIds(ReviewFinding finding) =>
        FindingExportDetails.GetCveIds(finding);

    public static bool HasActionableFix(ReviewFinding finding) =>
        !string.IsNullOrWhiteSpace(finding.OriginalFix) &&
        !ConnectSecureFixFormatter.IsPlaceholder(finding.OriginalFix);

    public static bool ShouldEnrich(ReviewFinding finding) =>
        GetCveIds(finding).Count > 0 &&
        string.IsNullOrWhiteSpace(finding.NvdEnrichment) &&
        (!HasActionableFix(finding) || CveReferenceHelper.IsCveOnlyProduct(finding.Product));

    public static bool ShouldAppendToExports(ReviewFinding finding) =>
        !string.IsNullOrWhiteSpace(finding.NvdEnrichment) && !HasActionableFix(finding);

    public static string AppendNvdContext(string baseText, ReviewFinding finding)
    {
        if (!ShouldAppendToExports(finding))
            return baseText;

        if (string.IsNullOrWhiteSpace(baseText))
            return finding.NvdEnrichment;

        if (baseText.Contains(finding.NvdEnrichment, StringComparison.Ordinal))
            return baseText;

        return baseText.TrimEnd() +
               Environment.NewLine +
               Environment.NewLine +
               "NVD / advisory context:" +
               Environment.NewLine +
               finding.NvdEnrichment;
    }
}
