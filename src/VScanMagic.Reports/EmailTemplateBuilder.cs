using VScanMagic.Core.Models;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class EmailTemplateBuilder
{
    private const string BakedInRmitNote =
        "Note: Remediation tickets have been generated for items covered under your RMIT+ agreement. Third-party items not covered under the agreement will not be remediated unless we discuss them and a quote has been generated. To schedule a discussion, please use the scheduling link above.";

    public static string Build(ReviewSession session, VScanMagicTemplates templates) =>
        Build(session, templates.ResolveEmailTemplate(session.IsRmitPlus), session.IsRmitPlus);

    public static string Build(ReviewSession session, EmailTemplateSettings template, bool isRmitPlus = false)
    {
        isRmitPlus = session.IsRmitPlus || isRmitPlus;
        var now = DateTime.Now;
        var quarter = (now.Month - 1) / 3 + 1;
        var exportCount = ReviewSessionRanker.GetExportFindings(session).Count;
        var topNLabel = session.ExportTopN <= 0 ? "Top"
            : session.ExportTopN == 10 && exportCount == 10 ? "Top Ten"
            : $"Top {exportCount}";

        var noteText = isRmitPlus
            ? "Note: Remediation tickets have been generated for items covered under your RMIT+ agreement. Third-party items not covered under the agreement will not be remediated unless we discuss them and a quote has been generated. To schedule a discussion, please use the scheduling link above."
            : "Note: No remediation will begin without your approval. To schedule a discussion, please use the scheduling link above.";

        var greeting = now.Hour switch
        {
            < 12 => "morning",
            < 17 => "afternoon",
            _ => "evening"
        };

        var subject = template.SubjectFormat
            .Replace("{Year}", now.Year.ToString(), StringComparison.Ordinal)
            .Replace("{Quarter}", quarter.ToString(), StringComparison.Ordinal);

        var body = template.Body
            .Replace("{Year}", now.Year.ToString(), StringComparison.Ordinal)
            .Replace("{Quarter}", quarter.ToString(), StringComparison.Ordinal)
            .Replace("{Greeting}", greeting, StringComparison.Ordinal)
            .Replace("{NoteText}", noteText, StringComparison.Ordinal)
            .Replace("{PreparedBy}", session.Presenter, StringComparison.Ordinal)
            .Replace("{TopNLabel}", topNLabel, StringComparison.Ordinal);

        if (!isRmitPlus && body.Contains(BakedInRmitNote, StringComparison.Ordinal))
            body = body.Replace(BakedInRmitNote, noteText, StringComparison.Ordinal);

        if (!body.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
            body = $"Subject: {subject}{Environment.NewLine}{Environment.NewLine}{body}";

        return NormalizeEmailSpacing(body);
    }

    public static (string Subject, string Body) SplitSubjectAndBody(string emailContent)
    {
        var normalized = emailContent.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        if (lines.Length > 0 && lines[0].StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
        {
            var subject = lines[0]["Subject:".Length..].Trim();
            var body = string.Join(Environment.NewLine, lines.Skip(1)).TrimStart('\n');
            return (subject, body);
        }

        return ("", normalized);
    }

    private static string NormalizeEmailSpacing(string text) =>
        string.Join(Environment.NewLine,
            text.Replace("\r\n", "\n").Split('\n')
                .Select(line => line.TrimEnd())
                .ToArray()).Trim();
}
