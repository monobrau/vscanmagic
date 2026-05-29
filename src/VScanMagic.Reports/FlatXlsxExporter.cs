using ClosedXML.Excel;
using VScanMagic.Review;
using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Reports;

public sealed class FlatXlsxExporter
{
    public void Export(ReviewSession session, string outputPath)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Top Findings");

        var headers = new[]
        {
            "Rank", "Source", "Product", "Status", "Risk Score", "EPSS", "Avg CVSS",
            "Critical", "High", "Medium", "Low", "Vuln Count",
            "Affected Hosts",
            "Revised Remediation", "Meeting Notes", "Tasks"
        };

        for (var c = 0; c < headers.Length; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        var row = 2;
        foreach (var f in ReviewSessionRanker.GetExportFindings(session))
        {
            ws.Cell(row, 1).Value = f.Rank;
            ws.Cell(row, 2).Value = f.Source;
            ws.Cell(row, 3).Value = ExcelCellText.Truncate(f.Product);
            ws.Cell(row, 4).Value = f.Status.ToString();
            ws.Cell(row, 5).Value = f.RiskScore;
            ws.Cell(row, 6).Value = f.Epss;
            ws.Cell(row, 7).Value = f.AvgCvss;
            ws.Cell(row, 8).Value = f.Critical;
            ws.Cell(row, 9).Value = f.High;
            ws.Cell(row, 10).Value = f.Medium;
            ws.Cell(row, 11).Value = f.Low;
            ws.Cell(row, 12).Value = FindingExportDetails.GetIncludedVulnCount(f);
            ws.Cell(row, 13).Value = ExcelCellText.FormatAffectedSystemsForExcel(f);
            ws.Cell(row, 14).Value = ExcelCellText.Truncate(FindingExportDetails.GetRemediationText(f));
            ws.Cell(row, 15).Value = ExcelCellText.Truncate(f.MeetingNotes);
            ws.Cell(row, 16).Value = ExcelCellText.Truncate(string.Join("; ", f.Tasks.Select(t => t.Text)));
            row++;
        }

        AddAffectedSystemsSheet(workbook, session);
        ws.Columns().AdjustToContents();
        workbook.SaveAs(outputPath);
    }

    private static void AddAffectedSystemsSheet(XLWorkbook workbook, ReviewSession session)
    {
        var ws = workbook.Worksheets.Add("Affected Systems");
        var headers = new[] { "Rank", "Product", "Host", "IP", "User", "Vuln Count" };
        for (var c = 0; c < headers.Length; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        var row = 2;
        foreach (var f in ReviewSessionRanker.GetExportFindings(session))
        {
            foreach (var system in FindingExportDetails.IncludedSystems(f))
            {
                ws.Cell(row, 1).Value = f.Rank;
                ws.Cell(row, 2).Value = ExcelCellText.Truncate(f.Product);
                ws.Cell(row, 3).Value = ExcelCellText.Truncate(system.HostName);
                ws.Cell(row, 4).Value = ExcelCellText.Truncate(system.Ip);
                ws.Cell(row, 5).Value = ExcelCellText.Truncate(system.Username);
                ws.Cell(row, 6).Value = system.VulnCount;
                row++;
            }
        }

        ws.Columns().AdjustToContents();
    }

    public void ExportSummaryFromRecords(IReadOnlyList<Core.Models.VulnerabilityRecord> records, string outputPath)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("All Findings");
        var headers = new[] { "Host", "IP", "Product", "Source", "Critical", "High", "Medium", "Low", "Count", "EPSS", "CVE", "Fix" };
        for (var c = 0; c < headers.Length; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        var row = 2;
        foreach (var r in records)
        {
            ws.Cell(row, 1).Value = ExcelCellText.Truncate(r.HostName);
            ws.Cell(row, 2).Value = ExcelCellText.Truncate(r.Ip);
            ws.Cell(row, 3).Value = ExcelCellText.Truncate(r.Product);
            ws.Cell(row, 4).Value = ExcelCellText.Truncate(r.Source);
            ws.Cell(row, 5).Value = r.Critical;
            ws.Cell(row, 6).Value = r.High;
            ws.Cell(row, 7).Value = r.Medium;
            ws.Cell(row, 8).Value = r.Low;
            ws.Cell(row, 9).Value = r.VulnerabilityCount;
            ws.Cell(row, 10).Value = r.EpssScore;
            ws.Cell(row, 11).Value = ExcelCellText.Truncate(r.Cve);
            ws.Cell(row, 12).Value = ExcelCellText.Truncate(ConnectSecureFixFormatter.ToReadableText(r.Fix));
            row++;
        }

        ws.Columns().AdjustToContents();
        workbook.SaveAs(outputPath);
    }
}
