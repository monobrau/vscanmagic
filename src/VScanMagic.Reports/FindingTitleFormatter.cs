using VScanMagic.Core.Risk;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class FindingTitleFormatter
{
    public static string FormatTicketSubject(ReviewFinding finding, bool isRmitPlus)
    {
        var suffix = CveExportFormatter.UsesCveExportTreatment(finding)
            ? CveExportFormatter.TicketSubjectSuffix
            : ProductTypeSuffixHelper.GetSuffix(finding.Product, isRmitPlus);
        var subject = $"Vulnerability Remediation - {finding.Product}{suffix}";

        if (isRmitPlus)
        {
            var modifier = TimeEstimateModifierHelper.GetModifierTextForSubject(
                finding.AfterHours, finding.TicketGenerated, finding.ThirdParty);
            if (!string.IsNullOrWhiteSpace(modifier))
                subject += modifier;
            if (finding.AfterHours)
                subject = $"After Hours - {subject}";
        }

        return subject;
    }

    public static string FormatDocxHeading(ReviewFinding finding, int rank, bool isRmitPlus)
    {
        var title = $"{rank}. {finding.Product}";

        if (isRmitPlus)
        {
            var modifier = TimeEstimateModifierHelper.GetModifierText(
                finding.AfterHours, finding.TicketGenerated, finding.ThirdParty);
            if (!string.IsNullOrWhiteSpace(modifier))
                title += modifier;
            if (finding.AfterHours)
                title = $"After Hours - {title}";
        }
        else if (CveExportFormatter.UsesCveExportTreatment(finding))
        {
            title += CveExportFormatter.TicketSubjectSuffix;
        }
        else
        {
            title += ProductTypeSuffixHelper.GetSuffix(finding.Product, false);
        }

        return title;
    }
}
