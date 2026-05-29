using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class PdfReviewExporter(RemediationRuleService remediationRules)
{
    static PdfReviewExporter()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    public void Export(ReviewSession session, string outputPath)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var findings = ReviewSessionRanker.GetExportFindings(session);
        var maxRisk = findings.Count > 0 ? findings.Max(f => f.RiskScore) : 10;
        var topLabel = ReviewExportLabels.GetTopNLabel(session);

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Margin(40);
                page.DefaultTextStyle(x => x.FontSize(10));
                page.Header().Text($"{topLabel} Report - {session.ClientName}").SemiBold().FontSize(16);
                page.Content().Column(col =>
                {
                    col.Item().Text($"Scan Date: {session.ScanDate}");
                    col.Item().Text($"Prepared by: {session.Presenter}");
                    col.Item().PaddingVertical(10).LineHorizontal(1);
                    col.Item().Text("Executive Summary").Bold().FontSize(14);
                    col.Item().Text(
                        $"Vulnerability assessment for {session.ClientName}. This read-only copy reflects review session state.");
                    col.Item().PaddingVertical(10);

                    for (var i = 0; i < findings.Count; i++)
                    {
                        var f = findings[i];
                        var (bg, _, name) = RiskScoreCalculator.GetRiskColor(f.RiskScore, maxRisk);
                        col.Item().Background($"#{bg}").Padding(6).Column(inner =>
                        {
                            inner.Item().Text($"{i + 1}. {f.Product} — {name} ({f.Status})").Bold();
                            inner.Item().Text($"Risk: {f.RiskScore:N2} | EPSS: {f.Epss:N4} | Vulns: {FindingExportDetails.GetIncludedVulnCount(f)}");
                        });

                        var systems = FindingExportDetails.IncludedSystems(f);
                        if (systems.Count > 0)
                        {
                            col.Item().PaddingTop(4).Text("Affected systems:").Bold();
                            foreach (var system in systems)
                                col.Item().Text($"• {FindingExportDetails.FormatAffectedSystem(system)}");
                        }

                        var cveReferences = CveExportFormatter.FormatReferencesSection(f);
                        if (!string.IsNullOrWhiteSpace(cveReferences))
                        {
                            col.Item().PaddingTop(4).Text("CVE references:").Bold();
                            col.Item().Text(cveReferences);
                        }

                        col.Item().PaddingTop(4).Text("Remediation:").Bold();
                        col.Item().PaddingBottom(4).Text(
                            FindingRemediationExport.GetWordRemediationText(f, remediationRules));
                        var connectSecureSolution = FindingRemediationExport.GetConnectSecureSolution(f);
                        if (connectSecureSolution is not null)
                        {
                            col.Item().PaddingTop(4).Text("ConnectSecure Solution:").Bold();
                            col.Item().Text(connectSecureSolution);
                        }

                        if (f.Tasks.Count > 0)
                        {
                            col.Item().PaddingTop(4).Text("Tasks:").Bold();
                            foreach (var t in f.Tasks)
                                col.Item().Text($"• {t.Text}");
                        }
                        col.Item().PaddingVertical(6).LineHorizontal(0.5f);
                    }
                });
                page.Footer().AlignCenter().Text(t =>
                {
                    t.Span("VScanMagic — Read-only client copy — Page ");
                    t.CurrentPageNumber();
                });
            });
        }).GeneratePdf(outputPath);
    }
}
