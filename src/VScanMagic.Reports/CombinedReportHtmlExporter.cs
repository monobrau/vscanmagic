using System.Net;
using System.Text;
using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class CombinedReportHtmlExporter(TemplatesService templatesService, RemediationRuleService remediationRules)
{
    public void Export(ReviewSession session, string outputPath, string? reportsPathPartial = null)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var templates = templatesService.Load();
        var html = BuildHtml(session, templates, reportsPathPartial);
        File.WriteAllText(outputPath, html, Encoding.UTF8);
    }

    public string BuildHtml(ReviewSession session, Core.Models.VScanMagicTemplates? templates = null, string? reportsPathPartial = null)
    {
        templates ??= templatesService.Load();
        var sections = TicketInstructionBuilder.BuildSections(session, remediationRules, reportsPathPartial);
        var ticketInstructionsHtml = BuildTicketInstructionsPanel(sections);
        var emailContent = EmailTemplateBuilder.Build(session, templates);
        var (emailSubject, _) = EmailTemplateBuilder.SplitSubjectAndBody(emailContent);
        var emailPanelHtml = BuildCopyPanel(
            panelId: "email",
            tabLabel: "Email Template",
            copyButtons: """
                <button type="button" class="copy-btn" onclick="copyEmailSubject()">Copy Subject</button>
                <button type="button" class="copy-btn" onclick="copyEmailBody()">Copy Body</button>
                """,
            contentId: "email-content",
            content: emailContent,
            useTextarea: true,
            dataAttributes: $"data-email-subject=\"{Html(emailSubject)}\"");

        var ticketNotes = TicketNotesBuilder.Build(session, templates.TicketNotes, session.IsRmitPlus);
        var timeEstimate = TimeEstimateBuilder.Build(session, remediationRules);
        var notesPanelHtml = BuildCopyPanel(
            panelId: "notes",
            tabLabel: "Ticket Notes",
            copyButtons: """<button type="button" class="copy-btn" onclick="copyTicketNotes()">Copy Ticket Notes</button>""",
            contentId: "ticket-notes-content",
            content: ticketNotes,
            useTextarea: false);

        var timePanelHtml = BuildCopyPanel(
            panelId: "time",
            tabLabel: "Time Estimate",
            copyButtons: """<button type="button" class="copy-btn" onclick="copyTimeEstimate()">Copy Time Estimate</button>""",
            contentId: "time-estimate-content",
            content: timeEstimate,
            useTextarea: false);

        var company = Html(session.ClientName);
        var generated = Html(DateTime.Now.ToString("MMMM d, yyyy"));
        var generatedAt = Html(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        var scanDate = Html(session.ScanDate);
        var titleTime = Html(DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
        var count = sections.Count;
        var headerTitle = count == 10 ? "TOP 10 VULNERABILITIES"
            : count > 0 ? $"TOP {count} VULNERABILITIES"
            : "VULNERABILITY REMEDIATIONS";

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset=\"UTF-8\">");
        sb.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
        sb.AppendLine($"  <title>Vulnerability Report - {company} - {titleTime}</title>");
        sb.AppendLine("  <style>");
        sb.AppendLine(HtmlStyles);
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("  <div class=\"report-header\">");
        sb.AppendLine($"    <h1 class=\"company-name\">{company}</h1>");
        sb.AppendLine($"    <div class=\"report-meta\">Vulnerability Report - {generated} | Scan date: {scanDate}</div>");
        sb.AppendLine("  </div>");
        sb.AppendLine("  <div class=\"tab-bar\">");
        sb.AppendLine("    <button class=\"tab-btn active\" data-tab=\"ticket\">Ticket Instructions</button>");
        sb.AppendLine("    <button class=\"tab-btn\" data-tab=\"email\">Email Template</button>");
        sb.AppendLine("    <button class=\"tab-btn\" data-tab=\"time\">Time Estimate</button>");
        sb.AppendLine("    <button class=\"tab-btn\" data-tab=\"notes\">Ticket Notes (Manage PSA)</button>");
        sb.AppendLine("  </div>");
        sb.AppendLine("  <div id=\"panel-ticket\" class=\"tab-panel active\">");
        sb.AppendLine("    <div class=\"header\">");
        sb.AppendLine($"      <h1>{Html(headerTitle)} - TICKET INSTRUCTIONS</h1>");
        sb.AppendLine($"      <div class=\"meta\">Generated: {generatedAt}</div>");
        sb.AppendLine("      <div class=\"key\">Use <strong>Copy Subject</strong> / <strong>Copy Section</strong> to paste into Manage PSA tickets, quotes, or technician assignments.</div>");
        sb.AppendLine("    </div>");
        sb.Append(ticketInstructionsHtml);
        sb.AppendLine("  </div>");
        sb.Append(emailPanelHtml);
        sb.Append(timePanelHtml);
        sb.Append(notesPanelHtml);
        sb.AppendLine(CopyScript);
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    private const string HtmlStyles = """
        body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .report-header { margin-bottom: 20px; padding: 16px 20px; background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .report-header .company-name { font-size: 24px; font-weight: bold; color: #1a1a1a; margin: 0; }
        .report-header .report-meta { color: #666; font-size: 13px; margin-top: 6px; }
        .tab-bar { display: flex; gap: 4px; margin-bottom: 16px; border-bottom: 2px solid #ddd; flex-wrap: wrap; }
        .tab-btn { padding: 10px 20px; cursor: pointer; background: #e9ecef; border: none; border-radius: 4px 4px 0 0; font-size: 14px; }
        .tab-btn:hover { background: #dee2e6; }
        .tab-btn.active { background: #fff; border: 1px solid #ddd; border-bottom: 2px solid #fff; margin-bottom: -2px; font-weight: bold; }
        .tab-panel { display: none; }
        .tab-panel.active { display: block; }
        .tab-actions { margin-bottom: 12px; }
        .tab-actions .copy-btn, .section-actions button { margin-right: 8px; padding: 8px 16px; cursor: pointer; background: #0066cc; color: #fff; border: none; border-radius: 4px; font-size: 13px; }
        .tab-actions .copy-btn:hover, .section-actions button:hover { background: #0052a3; }
        .tab-pre, .section-text { white-space: pre-wrap; overflow-x: auto; font-family: Consolas, monospace; font-size: 13px; line-height: 1.5; margin: 0; min-width: min(80ch, 100%); word-break: break-word; }
        .tab-pre { padding: 20px; background: #fff; border-radius: 8px; border: 1px solid #ddd; }
        textarea.tab-pre.email-template-text { display: block; width: 100%; box-sizing: border-box; min-height: 22rem; resize: vertical; border: 1px solid #ddd; overflow-wrap: anywhere; }
        .vuln-section { background: #fff; padding: 20px; margin-bottom: 24px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .vuln-section h2 { margin: 0 0 12px 0; font-size: 18px; color: #333; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
        .subject-line { font-size: 13px; color: #333; margin: 8px 0 12px 0; }
        .section-actions { margin-bottom: 0; }
        .section-actions .toggle-btn { background: #6c757d; }
        .section-actions .toggle-btn:hover { background: #5a6268; }
        .section-content { margin-top: 16px; padding-top: 16px; border-top: 1px solid #eee; }
        .vuln-section.collapsed .section-content { display: none; }
        .header { margin-bottom: 20px; }
        .header h1 { margin: 0; font-size: 20px; }
        .header .meta { color: #666; font-size: 12px; margin-top: 4px; }
        .key { margin-top: 12px; padding: 8px 12px; background: #fff3cd; border-radius: 4px; font-size: 12px; }
        """;

    private static string BuildTicketInstructionsPanel(IReadOnlyList<TicketInstructionSection> sections)
    {
        if (sections.Count == 0)
            return "<p>No export findings in this session.</p>";

        var sb = new StringBuilder();
        foreach (var section in sections)
        {
            var subject = Html(section.Subject);
            var bodyHtml = Html(section.BodyText).Replace("\n", "<br>\n", StringComparison.Ordinal);
            sb.AppendLine($"""
                <section id="{section.SectionId}" class="vuln-section collapsed">
                  <div class="section-header">
                    <h2>Vulnerability #{section.Number}</h2>
                    <div class="subject-line">{subject}</div>
                    <div class="section-actions">
                      <button type="button" onclick="copySubject(this)" data-subject="{subject}">Copy Subject</button>
                      <button type="button" onclick="copySection('{section.SectionId}')">Copy Section</button>
                      <button type="button" class="toggle-btn" onclick="toggleSection('{section.SectionId}')">+ Show details</button>
                    </div>
                  </div>
                  <div class="section-content">
                    <pre class="section-text">{bodyHtml}</pre>
                  </div>
                </section>
                """);
        }

        return sb.ToString();
    }

    private static string BuildCopyPanel(
        string panelId,
        string tabLabel,
        string copyButtons,
        string contentId,
        string content,
        bool useTextarea,
        string dataAttributes = "")
    {
        var encoded = Html(content);
        var inner = useTextarea
            ? $"""<textarea id="{contentId}" class="tab-pre email-template-text" readonly spellcheck="false">{encoded}</textarea>"""
            : $"""<pre id="{contentId}" class="tab-pre">{encoded}</pre>""";

        return $"""
            <div id="panel-{panelId}" class="tab-panel" {dataAttributes}>
              <div class="tab-actions">{copyButtons}</div>
              {inner}
            </div>
            """;
    }

    private static string Html(string? value) =>
        WebUtility.HtmlEncode(value ?? "");

    private const string CopyScript = """
      <script>
        document.querySelectorAll('.tab-btn').forEach(function(btn) {
          btn.addEventListener('click', function() {
            var tab = this.getAttribute('data-tab');
            document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
            document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('active'); });
            this.classList.add('active');
            var panel = document.getElementById('panel-' + tab);
            if (panel) panel.classList.add('active');
          });
        });
        function copyText(text) {
          if (!text) return;
          navigator.clipboard.writeText(text).catch(function() { prompt('Copy this:', text); });
        }
        function copySubject(btn) {
          copyText(btn.getAttribute('data-subject'));
        }
        function copySection(id) {
          var el = document.getElementById(id);
          var pre = el ? el.querySelector('.section-text') : null;
          if (!pre) return;
          copyText(pre.innerText || pre.textContent);
        }
        function toggleSection(id) {
          var el = document.getElementById(id);
          var btn = el ? el.querySelector('.toggle-btn') : null;
          if (!el || !btn) return;
          if (el.classList.contains('collapsed')) {
            el.classList.remove('collapsed');
            btn.textContent = '- Hide details';
          } else {
            el.classList.add('collapsed');
            btn.textContent = '+ Show details';
          }
        }
        function copyEmailSubject() {
          var panel = document.getElementById('panel-email');
          var subject = panel ? panel.getAttribute('data-email-subject') : '';
          copyText(subject);
        }
        function copyEmailBody() {
          var ta = document.getElementById('email-content');
          if (!ta || typeof ta.value !== 'string') return;
          var full = ta.value.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
          var lines = full.split('\n');
          var body = full;
          if (lines.length > 0 && /^Subject:\s*/i.test(lines[0])) {
            body = lines.slice(1).join('\n').replace(/^\n+/, '');
          }
          body = body.replace(/\n/g, '\r\n');
          copyText(body);
        }
        function copyTicketNotes() {
          var pre = document.getElementById('ticket-notes-content');
          if (!pre) return;
          copyText(pre.innerText || pre.textContent);
        }
        function copyTimeEstimate() {
          var pre = document.getElementById('time-estimate-content');
          if (!pre) return;
          copyText(pre.innerText || pre.textContent);
        }
      </script>
      """;
}
