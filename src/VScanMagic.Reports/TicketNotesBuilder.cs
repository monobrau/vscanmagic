using VScanMagic.Core.Models;
using VScanMagic.Core.Risk;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class TicketNotesBuilder
{
    public static string Build(ReviewSession session, TicketNotesTemplateSettings template, bool isRmitPlus = false)
    {
        isRmitPlus = session.IsRmitPlus || isRmitPlus;
        var topLabel = ReviewExportLabels.GetTopNLabel(session);

        var reportStepLine = $"Produced {topLabel} vulnerabilities docx report";
        var stepsBefore = template.StepsBeforeTickets.Replace("{ReportStepLine}", reportStepLine, StringComparison.Ordinal);

        var ticketLines = new List<string>();
        if (isRmitPlus)
        {
            foreach (var finding in ReviewSessionRanker.GetExportFindings(session))
            {
                if (TimeEstimateModifierHelper.IsTicketGenerated(
                        finding.AfterHours, finding.TicketGenerated, finding.ThirdParty))
                {
                    ticketLines.Add(FormatTicketNoteLine(finding, isRmitPlus));
                }
            }
        }

        var stepsText = ticketLines.Count > 0
            ? $"{stepsBefore.TrimEnd()}{Environment.NewLine}{string.Join(Environment.NewLine, ticketLines)}{Environment.NewLine}{template.StepsAfterTickets.Trim()}"
            : $"{stepsBefore.TrimEnd()}{Environment.NewLine}{template.StepsAfterTickets.Trim()}";

        return $"""
            Steps performed

            {stepsText}

            {template.ResolvedQuestion}

            {template.ResolvedAnswer}

            {template.NextStepsQuestion}

            {template.NextStepsText}
            """;
    }

    internal static string FormatTicketNoteLine(ReviewFinding finding, bool isRmitPlus)
    {
        var subject = FindingTitleFormatter.FormatTicketSubject(finding, isRmitPlus);
        if (!string.IsNullOrWhiteSpace(finding.ManageTicketNumber))
        {
            var statusSuffix = string.IsNullOrWhiteSpace(finding.ManageTicketStatus)
                ? ""
                : $" ({finding.ManageTicketStatus})";
            return $"- Ticket #{finding.ManageTicketNumber}{statusSuffix}: {subject}";
        }

        return $"- Ticket created: {subject}";
    }
}
