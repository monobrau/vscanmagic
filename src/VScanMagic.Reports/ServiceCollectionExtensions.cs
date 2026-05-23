using Microsoft.Extensions.DependencyInjection;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;

namespace VScanMagic.Reports;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicReports(this IServiceCollection services)
    {
        services.AddSingleton<DocxReviewExporter>();
        services.AddSingleton<PdfReviewExporter>();
        services.AddSingleton<TicketExporter>();
        services.AddSingleton<EmailExporter>();
        services.AddSingleton<FlatXlsxExporter>();
        services.AddSingleton<CombinedReportHtmlExporter>();
        services.AddSingleton<HostVulnerabilityReportExporter>();
        return services;
    }
}

public sealed record SessionExportOptions(
    bool IncludeHostVulnerabilityCounts = false,
    IReadOnlyList<Core.Models.HostVulnerabilitySummary>? HostCounts = null);

public sealed class ExportOrchestrator(
    DocxReviewExporter docx,
    PdfReviewExporter pdf,
    TicketExporter ticket,
    EmailExporter email,
    FlatXlsxExporter xlsx,
    CombinedReportHtmlExporter html,
    HostVulnerabilityReportExporter hostCountsExporter,
    TemplatesService templatesService,
    RemediationRuleService remediationRules)
{
    public ExportResult ExportAll(
        VScanMagic.Review.Models.ReviewSession session,
        ReportOutputLayout layout,
        SessionExportOptions? options = null)
    {
        options ??= new SessionExportOptions();
        Directory.CreateDirectory(layout.OutputDirectory);
        Directory.CreateDirectory(layout.TextOutputDirectory);

        var stamp = ReportPathResolver.GetReportTimestamp();
        var companyName = string.IsNullOrWhiteSpace(session.ClientName) ? "Client" : session.ClientName.Trim();
        var reportTitle = GetReportTitle(session);

        var docxPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.OutputDirectory, companyName, $" {reportTitle}_{stamp}", "docx");
        var pdfPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.OutputDirectory, companyName, $" Top Ten Review (Client)_{stamp}", "pdf");
        var ticketPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.TextOutputDirectory, companyName, $" Ticket Instructions_{stamp}", "txt");
        var htmlPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.TextOutputDirectory, companyName, $" Report_{stamp}", "html");
        var notesPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.TextOutputDirectory, companyName, $" Ticket Notes_{stamp}", "txt");
        var timeEstimatePath = ReportPathResolver.GetSafeReportOutputPath(
            layout.TextOutputDirectory, companyName, $" Time Estimate_{stamp}", "txt");
        var xlsxPath = ReportPathResolver.GetSafeReportOutputPath(
            layout.OutputDirectory, companyName, $" Top Ten Data_{stamp}", "xlsx");

        docx.Export(session, docxPath);
        pdf.Export(session, pdfPath);
        ticket.ExportToFile(session, ticketPath, layout.ReportsPathPartial);
        html.Export(session, htmlPath, layout.ReportsPathPartial);

        var (emailText, emailEml) = email.Export(session, layout.TextOutputDirectory, companyName, stamp);
        xlsx.Export(session, xlsxPath);

        var ticketNotes = TicketNotesBuilder.Build(session, templatesService.Load().TicketNotes, session.IsRmitPlus);
        File.WriteAllText(notesPath, ticketNotes);

        var timeEstimateText = TimeEstimateBuilder.Build(session, remediationRules);
        File.WriteAllText(timeEstimatePath, timeEstimateText);

        string? hostCountsPdf = null;
        string? hostCountsXlsx = null;
        if (options.IncludeHostVulnerabilityCounts && options.HostCounts is { Count: > 0 })
        {
            var hostRequest = new HostVulnerabilityReportRequest(
                session.ClientName,
                options.HostCounts,
                session.ScanDate,
                string.IsNullOrWhiteSpace(session.SourceFilePath) ? null : Path.GetFileName(session.SourceFilePath));
            var hostResult = hostCountsExporter.Export(hostRequest, layout.OutputDirectory, companyName, stamp, includePdf: true, includeXlsx: true);
            hostCountsPdf = hostResult.PdfPath;
            hostCountsXlsx = hostResult.XlsxPath;
        }

        return new ExportResult(
            layout.OutputDirectory,
            docxPath, pdfPath, ticketPath, htmlPath, notesPath, timeEstimatePath, emailText, emailEml, xlsxPath,
            hostCountsPdf, hostCountsXlsx);
    }

    private static string GetReportTitle(VScanMagic.Review.Models.ReviewSession session)
    {
        if (session.ExportTopN <= 0)
            return "Top Vulnerabilities Report";

        return session.ExportTopN == 10
            ? "Top Ten Vulnerabilities Report"
            : $"Top {session.ExportTopN} Vulnerabilities Report";
    }
}

public sealed record ExportResult(
    string OutputDirectory,
    string DocxPath,
    string PdfPath,
    string TicketPath,
    string HtmlPath,
    string TicketNotesPath,
    string TimeEstimatePath,
    string EmailTextPath,
    string EmailEmlPath,
    string XlsxPath,
    string? HostCountsPdfPath = null,
    string? HostCountsXlsxPath = null);
