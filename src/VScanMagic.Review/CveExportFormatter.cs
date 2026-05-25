using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

/// <summary>
/// Package 1 CVE export treatment: CVE-only product names, NVD URLs, identify-device steps (no generic * rule).
/// </summary>
public static class CveExportFormatter
{
    /// <summary>Ticket/deliverables subject suffix for CVE-only findings (registry/CVE-named rows).</summary>
    public const string TicketSubjectSuffix = " - Investigate and Resolve";

    private const string MinimalTicketSteps =
        "- Identify the affected software, firmware, or device on the listed host(s) (this finding is reported by CVE id only)\r\n" +
        "- Open each NVD link in CVE references and follow vendor security advisories for patches, firmware, or configuration mitigations\r\n" +
        "- Apply updates via RMM or scripting when available for the client; otherwise perform manual remediation per the vendor guidance\r\n" +
        "- Document the resolved product/version in the ticket when remediation is complete";

    /// <summary>CVE-only product rows (e.g. registry findings named CVE-2015-0240).</summary>
    public static bool UsesCveExportTreatment(ReviewFinding finding) =>
        CveReferenceHelper.IsCveOnlyProduct(finding.Product) &&
        FindingExportDetails.GetCveIds(finding).Count > 0;

    public static string FormatReferencesSection(ReviewFinding finding)
    {
        if (!UsesCveExportTreatment(finding))
            return "";

        var lines = new List<string>();
        foreach (var cveId in FindingExportDetails.GetCveIds(finding))
        {
            lines.Add(cveId);
            lines.Add(CveReferenceHelper.GetNvdDetailUrl(cveId));
        }

        return string.Join(Environment.NewLine, lines);
    }

    public static string GetMinimalRemediationSteps() => MinimalTicketSteps;

    public static string GetRemediationInstructions(
        ReviewFinding finding,
        RemediationRuleService remediationRules,
        bool forWord)
    {
        if (!UsesCveExportTreatment(finding))
        {
            if (FindingRemediationExport.IsRemediationEdited(finding))
                return CveEnrichmentPolicy.AppendNvdContext(FindingExportDetails.GetRemediationText(finding), finding);

            var guidance = remediationRules.GetGuidance(finding.Product, forWord);
            return CveEnrichmentPolicy.AppendNvdContext(guidance, finding);
        }

        if (FindingRemediationExport.IsRemediationEdited(finding))
            return FindingExportDetails.GetRemediationText(finding);

        return forWord
            ? MinimalTicketSteps.Replace("\r\n", Environment.NewLine)
            : MinimalTicketSteps;
    }
}
