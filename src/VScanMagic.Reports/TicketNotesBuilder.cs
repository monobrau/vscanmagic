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
        var exportCount = ReviewSessionRanker.GetExportFindings(session).Count;
        var topLabel = session.ExportTopN <= 0 ? "Top"
            : session.ExportTopN == 10 && exportCount == 10 ? "Top Ten"
            : $"Top {exportCount}";

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
                    ticketLines.Add($"- Ticket created for {finding.Product}");
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
}
