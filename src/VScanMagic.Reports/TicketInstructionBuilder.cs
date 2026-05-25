using System.Text;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed record TicketInstructionSection(int Number, string Subject, string BodyText, string SectionId);

public static class TicketInstructionBuilder
{
    private const string SuppressFooter =
        "Sometimes it will not be possible to remediate the vulnerability for business or technical reasons. Other times it will be a false positive detection. In the event of either case please reach out to someone on the Security team with your findings and we can suppress the vulnerability so it doesn't come up on future scans or remediations.";

    private const string UninstallFooter =
        "Uninstalling the software or removing/replacing the device is also a valid form of remediation when updating or patching is not feasible; the vulnerability will show as remediated on the next scan.";

    public static IReadOnlyList<TicketInstructionSection> BuildSections(
        ReviewSession session,
        RemediationRuleService remediationRules,
        string? reportsPathPartial = null)
    {
        var findings = ReviewSessionRanker.GetExportFindings(session);
        var sections = new List<TicketInstructionSection>(findings.Count);

        for (var i = 0; i < findings.Count; i++)
        {
            var finding = findings[i];
            var num = i + 1;
            sections.Add(new TicketInstructionSection(
                num,
                BuildSubject(finding, session.IsRmitPlus),
                BuildBodyText(finding, remediationRules, session.IsRmitPlus, reportsPathPartial),
                $"vuln-{num}"));
        }

        return sections;
    }

    public static string BuildPlainTextDocument(
        ReviewSession session,
        RemediationRuleService remediationRules,
        string? reportsPathPartial = null)
    {
        var findings = ReviewSessionRanker.GetExportFindings(session);
        var count = findings.Count;
        var headerTitle = count == 10 ? "TOP 10 VULNERABILITIES"
            : count > 0 ? $"TOP {count} VULNERABILITIES"
            : "VULNERABILITY REMEDIATIONS";

        var sb = new StringBuilder();
        sb.AppendLine(new string('=', 100));
        sb.AppendLine($"{headerTitle} - TICKET INSTRUCTIONS");
        sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine(new string('=', 100));

        foreach (var section in BuildSections(session, remediationRules, reportsPathPartial))
        {
            sb.AppendLine();
            sb.AppendLine(new string('-', 100));
            sb.AppendLine($"VULNERABILITY #{section.Number}");
            sb.AppendLine(new string('-', 100));
            sb.AppendLine();
            sb.AppendLine("TICKET SUBJECT:");
            sb.AppendLine(section.Subject);
            sb.AppendLine();
            sb.AppendLine(section.BodyText);
        }

        sb.AppendLine();
        sb.AppendLine(new string('=', 100));
        sb.AppendLine("END OF TICKET INSTRUCTIONS");
        sb.AppendLine(new string('=', 100));
        return NormalizeSpacing(sb.ToString());
    }

    public static string BuildSubject(ReviewFinding finding, bool isRmitPlus = false) =>
        FindingTitleFormatter.FormatTicketSubject(finding, isRmitPlus);

    public static string BuildBodyText(
        ReviewFinding finding,
        RemediationRuleService remediationRules,
        bool isRmitPlus = false,
        string? reportsPathPartial = null)
    {
        var systems = FindingExportDetails.IncludedSystems(finding);
        var sb = new StringBuilder();

        sb.AppendLine(BuildSubject(finding, isRmitPlus));
        sb.AppendLine();
        sb.AppendLine($"{"Product/System:".PadRight(25)}{finding.Product}");
        sb.AppendLine($"{"Risk Score:".PadRight(25)}{finding.RiskScore:N2}");
        sb.AppendLine($"{"EPSS Score:".PadRight(25)}{finding.Epss:N4}");
        sb.AppendLine($"{"Average CVSS:".PadRight(25)}{finding.AvgCvss:N2}");
        sb.AppendLine($"{"Total Vulnerabilities:".PadRight(25)}{finding.VulnCount}");
        sb.AppendLine($"{"Affected Systems Count:".PadRight(25)}{systems.Count}");

        var cveReferences = CveExportFormatter.FormatReferencesSection(finding);
        if (!string.IsNullOrWhiteSpace(cveReferences))
        {
            sb.AppendLine();
            sb.AppendLine("CVE references:");
            sb.AppendLine(cveReferences);
        }

        sb.AppendLine();
        sb.AppendLine("NOTE: This remediation can go to any available technician.");
        sb.AppendLine();
        sb.AppendLine("Affected Systems:");
        if (systems.Count == 0)
            sb.AppendLine("  (none included after review exclusions)");
        else
            sb.AppendLine(FindingExportDetails.FormatAffectedSystemsMultiline(finding));

        sb.AppendLine();
        sb.AppendLine("Remediation Instructions:");
        sb.AppendLine(FindingRemediationExport.GetTicketRemediationText(finding, remediationRules));

        var connectSecureSolution = FindingRemediationExport.GetConnectSecureSolution(finding);
        if (connectSecureSolution is not null)
        {
            sb.AppendLine();
            sb.AppendLine("ConnectSecure Solution:");
            sb.AppendLine(connectSecureSolution);
        }

        if (!string.IsNullOrWhiteSpace(finding.MeetingNotes))
        {
            sb.AppendLine();
            sb.AppendLine("Meeting Notes:");
            sb.AppendLine(finding.MeetingNotes);
        }

        if (finding.Tasks.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Tasks:");
            foreach (var task in finding.Tasks)
                sb.AppendLine($"- {task.Text}");
        }

        if (!string.IsNullOrWhiteSpace(reportsPathPartial))
        {
            sb.AppendLine();
            sb.AppendLine($"Reports location: ...\\{reportsPathPartial}");
        }

        sb.AppendLine();
        sb.AppendLine(UninstallFooter);
        sb.AppendLine();
        sb.AppendLine(SuppressFooter);

        return NormalizeSpacing(sb.ToString());
    }

    private static string NormalizeSpacing(string text)
    {
        var result = new List<string>();
        var blank = false;
        foreach (var line in text.Replace("\r\n", "\n").Split('\n'))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                if (!blank && result.Count > 0)
                {
                    result.Add("");
                    blank = true;
                }
            }
            else
            {
                result.Add(line.TrimEnd());
                blank = false;
            }
        }

        return string.Join(Environment.NewLine, result).Trim();
    }
}
