using System.Net;
using System.Text.RegularExpressions;
using VScanMagic.Core.Models;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public static class EmailTemplateBuilder
{
    private const string BakedInRmitNote =
        "Note: Remediation tickets have been generated for items covered under your RMIT+ agreement. Third-party items not covered under the agreement will not be remediated unless we discuss them and a quote has been generated. To schedule a discussion, please use the scheduling link above.";

    public const string SchedulingLinkLabel = "Schedule Time With Me";

    public static string Build(ReviewSession session, VScanMagicTemplates templates) =>
        Build(session, templates.ResolveEmailTemplate(session.IsRmitPlus), session.IsRmitPlus);

    public static string Build(ReviewSession session, VScanMagicTemplates templates, DeliverableLinks links) =>
        Build(session, templates.ResolveEmailTemplate(session.IsRmitPlus), session.IsRmitPlus, links);

    public static string Build(ReviewSession session, EmailTemplateSettings template, bool isRmitPlus = false) =>
        Build(session, template, isRmitPlus, new DeliverableLinks());

    public static string Build(
        ReviewSession session,
        EmailTemplateSettings template,
        bool isRmitPlus,
        DeliverableLinks links)
    {
        isRmitPlus = session.IsRmitPlus || isRmitPlus;
        var now = DateTime.Now;
        var quarter = (now.Month - 1) / 3 + 1;
        var topNLabel = GetTopNLabel(session);

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
            .Replace("{TopNLabel}", topNLabel, StringComparison.Ordinal)
            .Replace("{TopNReportLink}", links.TopNReportUrl, StringComparison.Ordinal)
            .Replace("{ReportsFolderLink}", links.ReportsFolderUrl, StringComparison.Ordinal)
            .Replace("{SchedulingLink}", links.SchedulingLinkUrl, StringComparison.Ordinal)
            .Replace("<link to top ten report from onedrive>", links.TopNReportUrl, StringComparison.OrdinalIgnoreCase)
            .Replace("<onedrive link to folder containing reports>", links.ReportsFolderUrl, StringComparison.OrdinalIgnoreCase)
            .Replace("<scheduling link>", links.SchedulingLinkUrl, StringComparison.OrdinalIgnoreCase);

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

    public static string GetTopNLabel(ReviewSession session)
    {
        var exportCount = ReviewSessionRanker.GetExportFindings(session).Count;
        return session.ExportTopN <= 0 ? "Top"
            : session.ExportTopN == 10 && exportCount == 10 ? "Top Ten"
            : $"Top {exportCount}";
    }

    public static string NormalizeDeliverableBodySpacing(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
            return "";

        var text = body.Replace("\r\n", "\n").Replace('\r', '\n');
        text = Regex.Replace(
            text,
            @"(?m)^Schedule time with me\s*\n+(?=https?://|Schedule time with me|Schedule Time With Me)",
            "",
            RegexOptions.IgnoreCase);
        text = Regex.Replace(
            text,
            @"(?m)^Schedule time with me\s*\n+Schedule time with me\s*(?=\n|$)",
            SchedulingLinkLabel,
            RegexOptions.IgnoreCase);
        text = CollapseInnerBulletBlanks(text);
        text = EnsureBlankLineAfterLinkLines(text);
        text = Regex.Replace(
            text,
            @"(?m)^(Not all vulnerabilities may be feasible[^\n]*)\n(?!\n)(?=Schedule time with me|Schedule Time With Me|https?://|Open Top|Open complete)",
            "$1\n\n",
            RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\n{3,}", "\n\n");
        return text.Trim();
    }

    private static string EnsureBlankLineAfterLinkLines(string text)
    {
        var lines = text.Split('\n');
        var output = new List<string>(lines.Length);
        for (var i = 0; i < lines.Length; i++)
        {
            output.Add(lines[i]);
            if (!IsLinkLine(lines[i]) || i + 1 >= lines.Length)
                continue;

            if (lines[i + 1].Length == 0)
                continue;

            output.Add("");
        }

        return string.Join('\n', output);
    }

    private static bool IsLinkLine(string line)
    {
        var trimmed = line.Trim();
        if (trimmed.Length == 0)
            return false;

        if (Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) &&
            (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
             uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            return true;

        if (trimmed.StartsWith("Open ", StringComparison.OrdinalIgnoreCase) &&
            trimmed.EndsWith(" report", StringComparison.OrdinalIgnoreCase))
            return true;

        if (trimmed.Equals("Open complete report package", StringComparison.OrdinalIgnoreCase))
            return true;

        return trimmed.Equals(SchedulingLinkLabel, StringComparison.OrdinalIgnoreCase) ||
               trimmed.Equals("Schedule time with me", StringComparison.OrdinalIgnoreCase);
    }

    private static string CollapseInnerBulletBlanks(string text)
    {
        var lines = text.Split('\n');
        var output = new List<string>(lines.Length);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            if (line.Length == 0 && output.Count > 0)
            {
                var previous = output[^1];
                var nextIndex = i + 1;
                while (nextIndex < lines.Length && lines[nextIndex].Length == 0)
                    nextIndex++;

                if (nextIndex < lines.Length)
                {
                    var next = lines[nextIndex];
                    if (IsBulletLine(previous) && IsBulletLine(next))
                    {
                        continue;
                    }
                }
            }

            output.Add(line);
        }

        return string.Join('\n', output);
    }

    private static bool IsBulletLine(string line) => line.TrimStart().StartsWith('•');

    public static string ApplyFriendlyLinkLabels(string body, DeliverableLinks links, string topNLabel)
    {
        if (string.IsNullOrWhiteSpace(body))
            return "";

        var replacements = GetTemplateLinkReplacements(
            links.TopNReportUrl,
            $"Open {topNLabel} report",
            links.ReportsFolderUrl,
            "Open complete report package",
            links.SchedulingLinkUrl,
            SchedulingLinkLabel);
        var text = ReplaceUrlLinesInDocumentOrder(body, replacements);
        text = NormalizeDeliverableBodySpacing(text);
        return text.Replace("\n", Environment.NewLine);
    }

    private static List<(string Url, string Label)> GetTemplateLinkReplacements(
        string topNUrl,
        string topNLabel,
        string folderUrl,
        string folderLabel,
        string schedulingUrl,
        string schedulingLabel)
    {
        var entries = new List<(string Url, string Label)>();
        if (!string.IsNullOrWhiteSpace(topNUrl))
            entries.Add((topNUrl.Trim(), topNLabel));
        if (!string.IsNullOrWhiteSpace(folderUrl))
            entries.Add((folderUrl.Trim(), folderLabel));
        if (!string.IsNullOrWhiteSpace(schedulingUrl))
            entries.Add((schedulingUrl.Trim(), schedulingLabel));

        return entries;
    }

    private static string ReplaceUrlLinesInDocumentOrder(
        string body,
        IReadOnlyList<(string Url, string Label)> orderedReplacements)
    {
        if (string.IsNullOrWhiteSpace(body) || orderedReplacements.Count == 0)
            return body;

        var pending = new Queue<(string Url, string Label)>(orderedReplacements);
        var lines = body.Split('\n');
        for (var i = 0; i < lines.Length && pending.Count > 0; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed.Length == 0)
                continue;

            if (!string.Equals(trimmed, pending.Peek().Url, StringComparison.OrdinalIgnoreCase))
                continue;

            lines[i] = pending.Dequeue().Label;
        }

        return string.Join('\n', lines);
    }

    public static (string PlainCopy, string HtmlBody) PrepareDeliverableCopy(
        string emailBody,
        DeliverableLinks links,
        string topNLabel)
    {
        var spaced = NormalizeDeliverableBodySpacing(emailBody);
        var plain = ApplyFriendlyLinkLabels(spaced, links, topNLabel);
        var html = BuildHtmlBody(spaced, links, topNLabel);
        return (plain, html);
    }

    public static string BuildHtmlBody(string plainBody, DeliverableLinks links, string topNLabel)
    {
        if (string.IsNullOrWhiteSpace(plainBody))
            return "";

        var normalized = NormalizeDeliverableBodySpacing(plainBody);
        var htmlLines = EmbedDeliverableLinkAnchors(normalized, links, topNLabel);
        return ConvertEncodedPlainTextToHtml(string.Join('\n', htmlLines));
    }

    private static List<string> EmbedDeliverableLinkAnchors(string normalizedPlain, DeliverableLinks links, string topNLabel)
    {
        var replacements = GetTemplateLinkReplacements(
                links.TopNReportUrl,
                $"Open {topNLabel} report",
                links.ReportsFolderUrl,
                "Open complete report package",
                links.SchedulingLinkUrl,
                SchedulingLinkLabel);

        var pending = new Queue<(string Url, string Label)>(replacements);
        var lines = normalizedPlain.Split('\n');
        var htmlLines = new List<string>(lines.Length);
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            string? anchor = null;
            if (pending.Count > 0 &&
                string.Equals(trimmed, pending.Peek().Url, StringComparison.OrdinalIgnoreCase))
            {
                var (url, label) = pending.Dequeue();
                anchor = $"<a href=\"{HtmlAttribute(url)}\">{WebUtility.HtmlEncode(label)}</a>";
            }

            htmlLines.Add(anchor ?? WebUtility.HtmlEncode(line));
        }

        return htmlLines;
    }

    public static string BuildHtmlBody(ReviewSession session, string plainBody, DeliverableLinks links) =>
        BuildHtmlBody(plainBody, links, GetTopNLabel(session));

    private static string ConvertEncodedPlainTextToHtml(string encodedBody)
    {
        var normalized = encodedBody.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        if (normalized.Length == 0)
            return "";

        const string paragraphStyle = "margin:0 0 12px 0;";
        const string spacerStyle = "margin:0 0 12px 0;mso-line-height-rule:exactly;line-height:14pt;font-size:14pt;";
        var lines = normalized.Split('\n');
        var html = new System.Text.StringBuilder(normalized.Length * 2);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            if (string.IsNullOrEmpty(line))
            {
                html.Append($"<p style=\"{spacerStyle}\">&nbsp;</p>");
                continue;
            }

            if (line.Contains("<a ", StringComparison.Ordinal) &&
                i + 1 < lines.Length &&
                lines[i + 1].Length > 0)
            {
                html.Append($"<p style=\"{paragraphStyle}\">{line}<br><br></p>");
                continue;
            }

            html.Append($"<p style=\"{paragraphStyle}\">{line}</p>");
        }

        return html.ToString();
    }

    private static string HtmlAttribute(string value) =>
        WebUtility.HtmlEncode(value);

    private static string NormalizeEmailSpacing(string text) =>
        string.Join(Environment.NewLine,
            text.Replace("\r\n", "\n").Split('\n')
                .Select(line => line.TrimEnd())
                .ToArray()).Trim();
}
