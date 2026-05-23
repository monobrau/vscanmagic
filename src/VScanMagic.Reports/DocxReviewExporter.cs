using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VScanMagic.Core.Risk;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class DocxReviewExporter
{
    public void Export(ReviewSession session, string outputPath)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        if (File.Exists(outputPath))
            File.Delete(outputPath);

        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        var findings = ReviewSessionRanker.GetExportFindings(session);
        var maxRisk = findings.Count > 0 ? findings.Max(f => f.RiskScore) : 10;
        var topLabel = session.ExportTopN <= 0 ? "Top" : (findings.Count == session.ExportTopN ? $"Top {session.ExportTopN}" : $"Top {findings.Count}");

        AppendParagraph(body, "Top Ten Vulnerabilities Report", bold: true, size: 32, center: true);
        AppendParagraph(body, session.ClientName, bold: true, size: 24, center: true);
        AppendParagraph(body, $"Scan Date: {session.ScanDate}", size: 22, center: true);
        AppendParagraph(body, $"Prepared by: {session.Presenter}", size: 20, center: true);
        AppendParagraph(body, "");
        AppendParagraph(body, "Executive Summary", bold: true, size: 28);
        AppendParagraph(body,
            $"This vulnerability assessment report summarizes the security posture of {session.ClientName} " +
            $"based on the vulnerability scan conducted on {session.ScanDate}. " +
            "Findings below reflect live review updates including agreed remediation and tasks.");
        AppendParagraph(body, "");

        AppendParagraph(body, $"{topLabel} Vulnerabilities by Risk Score", bold: true, size: 26);
        AppendSummaryTable(body, findings, maxRisk);
        AppendParagraph(body, "");

        AppendParagraph(body, "Detailed Findings and Remediation Guidance", bold: true, size: 26);
        for (var i = 0; i < findings.Count; i++)
        {
            var f = findings[i];
            AppendParagraph(body, $"{FindingTitleFormatter.FormatDocxHeading(f, i + 1, session.IsRmitPlus)} [{f.Status}]", bold: true, size: 24);
            AppendParagraph(body, $"Risk Score: {f.RiskScore:N2} | EPSS: {f.Epss:N4} | Avg CVSS: {f.AvgCvss:N2} | Total Vulns: {f.VulnCount}");
            AppendParagraph(body, "Affected Systems:", bold: true);
            AppendAffectedSystems(body, f);
            AppendParagraph(body, "Remediation Guidance:", bold: true);
            AppendParagraph(body, FindingRemediationExport.GetWordRemediationText(f));
            var connectSecureSolution = FindingRemediationExport.GetConnectSecureSolution(f);
            if (connectSecureSolution is not null)
            {
                AppendParagraph(body, "ConnectSecure Solution:", bold: true);
                AppendParagraph(body, connectSecureSolution);
            }
            if (!string.IsNullOrWhiteSpace(f.MeetingNotes))
            {
                AppendParagraph(body, "Meeting Notes:", bold: true);
                AppendParagraph(body, f.MeetingNotes);
            }
            if (f.Tasks.Count > 0)
            {
                AppendParagraph(body, "Tasks:", bold: true);
                foreach (var t in f.Tasks)
                {
                    var line = $"- {t.Text}";
                    if (!string.IsNullOrWhiteSpace(t.Owner)) line += $" (Owner: {t.Owner})";
                    if (t.DueDate.HasValue) line += $" (Due: {t.DueDate:yyyy-MM-dd})";
                    AppendParagraph(body, line);
                }
            }
            AppendParagraph(body, "");
        }

        mainPart.Document.Save();
    }

    private static void AppendAffectedSystems(Body body, ReviewFinding finding)
    {
        var systems = FindingExportDetails.IncludedSystems(finding);
        if (systems.Count == 0)
        {
            AppendParagraph(body, "(none included after review exclusions)");
            return;
        }

        var table = new Table(
            new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        table.Append(CreateRow(["Host", "IP", "User", "Vulns"], header: true));
        foreach (var system in systems)
        {
            table.Append(CreateRow([
                string.IsNullOrWhiteSpace(system.HostName) ? system.Ip : system.HostName,
                system.Ip,
                system.Username,
                system.VulnCount.ToString()
            ]));
        }

        body.Append(table);
        AppendParagraph(body, "");
    }

    private static void AppendSummaryTable(Body body, IReadOnlyList<ReviewFinding> findings, double maxRisk)
    {
        var table = new Table(
            new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        table.Append(CreateRow(["Rank", "Product", "Risk Score", "EPSS", "Hosts", "Status"], header: true));
        for (var i = 0; i < findings.Count; i++)
        {
            var f = findings[i];
            var (bg, fg, _) = RiskScoreCalculator.GetRiskColor(f.RiskScore, maxRisk);
            table.Append(CreateRow(
                [
                    (i + 1).ToString(),
                    f.Product,
                    f.RiskScore.ToString("N2"),
                    f.Epss.ToString("N4"),
                    FindingExportDetails.IncludedSystems(f).Count.ToString(),
                    f.Status.ToString()
                ],
                bgHex: bg, fgHex: fg));
        }

        body.Append(table);
    }

    private static TableRow CreateRow(string[] cells, bool header = false, string? bgHex = null, string? fgHex = null)
    {
        var row = new TableRow();
        foreach (var text in cells)
        {
            var cell = new TableCell(new Paragraph(new Run(new Text(text))));
            if (header)
            {
                cell.Append(new TableCellProperties(new Shading { Fill = "4472C4", Val = ShadingPatternValues.Clear }));
                cell.Descendants<Run>().First().RunProperties = new RunProperties(new Bold(), new Color { Val = "FFFFFF" });
            }
            else if (bgHex is not null)
            {
                cell.Append(new TableCellProperties(new Shading { Fill = bgHex, Val = ShadingPatternValues.Clear }));
                if (fgHex is not null)
                    cell.Descendants<Run>().First().RunProperties = new RunProperties(new Color { Val = fgHex });
            }
            row.Append(cell);
        }
        return row;
    }

    private static void AppendParagraph(Body body, string text, bool bold = false, int size = 22, bool center = false)
    {
        var props = new ParagraphProperties();
        if (center) props.Append(new Justification { Val = JustificationValues.Center });

        var runProps = new RunProperties();
        runProps.Append(new FontSize { Val = (size * 2).ToString() });
        if (bold) runProps.Append(new Bold());

        body.Append(new Paragraph(props, new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }
}
