using System.Text;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class TimeEstimateBuilder
{
    public static string Build(ReviewSession session, RemediationRuleService remediationRules)
    {
        var findings = ReviewSessionRanker.GetExportFindings(session);
        var sb = new StringBuilder();
        sb.AppendLine(new string('=', 100));
        sb.AppendLine("TIME ESTIMATE FOR VULNERABILITY REMEDIATION");
        sb.AppendLine(new string('=', 100));
        sb.AppendLine();
        sb.AppendLine("Remediation Time Estimates:");
        sb.AppendLine();

        decimal totalCovered = 0;
        decimal totalRequiringApproval = 0;
        decimal grandTotal = 0;

        for (var i = 0; i < findings.Count; i++)
        {
            var finding = findings[i];
            AppendFindingBlock(sb, finding, session.IsRmitPlus, remediationRules, ref totalCovered, ref totalRequiringApproval, ref grandTotal, i + 1);
        }

        sb.AppendLine();
        sb.AppendLine(new string('=', 100));
        sb.AppendLine("SUMMARY");
        sb.AppendLine(new string('=', 100));
        sb.AppendLine();

        if (session.IsRmitPlus)
        {
            sb.AppendLine($"Total Covered by Agreement: {totalCovered} hours");
            sb.AppendLine($"Total Requiring Approval: {totalRequiringApproval} hours");
            sb.AppendLine();
        }

        sb.AppendLine($"Grand Total: {grandTotal} hours");

        if (!session.IsRmitPlus)
        {
            sb.AppendLine();
            sb.AppendLine("Note: We will not begin remediation without your prior approval.");
        }

        return sb.ToString().TrimEnd();
    }

    private static void AppendFindingBlock(
        StringBuilder sb,
        ReviewFinding finding,
        bool isRmitPlus,
        RemediationRuleService remediationRules,
        ref decimal totalCovered,
        ref decimal totalRequiringApproval,
        ref decimal grandTotal,
        int rank)
    {
        sb.AppendLine($"{rank}. {finding.Product}");

        var hosts = FindingExportDetails.FormatAffectedSystemsCompactInline(finding);
        if (!string.IsNullOrWhiteSpace(hosts))
            sb.AppendLine($"   Affected Hostnames: {hosts}");

        var remediation = FindingRemediationExport.GetTimeEstimateRemediationText(finding, remediationRules);
        if (!string.IsNullOrWhiteSpace(remediation))
        {
            sb.AppendLine("   Remediation Guidance:");
            foreach (var line in remediation.Replace("\r\n", "\n").Split('\n'))
            {
                if (!string.IsNullOrWhiteSpace(line))
                    sb.AppendLine($"     {line.Trim()}");
            }
        }

        if (isRmitPlus)
            AppendRmitPlusEstimate(sb, finding, ref totalCovered, ref totalRequiringApproval, ref grandTotal);
        else
            AppendHourlyEstimate(sb, finding, ref grandTotal);

        sb.AppendLine();
    }

    private static void AppendRmitPlusEstimate(
        StringBuilder sb,
        ReviewFinding finding,
        ref decimal totalCovered,
        ref decimal totalRequiringApproval,
        ref decimal grandTotal)
    {
        var isThirdParty = finding.ThirdParty;
        var isTicketGenerated = TimeEstimateModifierHelper.IsTicketGenerated(
            finding.AfterHours, finding.TicketGenerated, isThirdParty);
        var requiresApproval = finding.AfterHours || isThirdParty;

        if (isTicketGenerated)
        {
            if (isThirdParty && finding.AfterHours)
            {
                sb.AppendLine("   After Hours: Yes");
                sb.AppendLine("   Ticket Generated: Yes (Covered by Agreement - Auto-generated)");
            }
            else
            {
                sb.AppendLine("   Ticket Generated: Yes (Covered by Agreement)");
            }

            sb.AppendLine($"   Estimated Time: {finding.TimeEstimateHours} hours - A remediation ticket has already been generated");
            sb.AppendLine("   Status: Covered by Agreement");
            totalCovered += finding.TimeEstimateHours;
        }
        else if (!requiresApproval)
        {
            sb.AppendLine($"   Estimated Time: {finding.TimeEstimateHours} hours - A remediation ticket has already been generated");
            sb.AppendLine("   Status: Covered by Agreement");
            totalCovered += finding.TimeEstimateHours;
        }
        else
        {
            if (finding.AfterHours)
            {
                sb.AppendLine("   After Hours: Yes");
                sb.AppendLine("   Estimated Time: N/A - A remediation ticket has already been generated");
            }
            else
            {
                sb.AppendLine($"   Estimated Time: {finding.TimeEstimateHours} hours");
            }

            sb.AppendLine("   Status: Requires Approval");
            if (!finding.AfterHours)
            {
                totalRequiringApproval += finding.TimeEstimateHours;
                grandTotal += finding.TimeEstimateHours;
            }
        }
    }

    private static void AppendHourlyEstimate(StringBuilder sb, ReviewFinding finding, ref decimal grandTotal)
    {
        if (finding.AfterHours)
        {
            sb.AppendLine("   After Hours: Yes");
            sb.AppendLine($"   Estimated Time: {finding.TimeEstimateHours} hours");
        }
        else
        {
            sb.AppendLine($"   Estimated Time: {finding.TimeEstimateHours} hours");
        }

        grandTotal += finding.TimeEstimateHours;
    }
}
