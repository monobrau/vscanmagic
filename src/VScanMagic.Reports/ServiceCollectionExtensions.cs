using Microsoft.Extensions.DependencyInjection;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;
using VScanMagic.Review;

namespace VScanMagic.Reports;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicReports(this IServiceCollection services)
    {
        services.AddSingleton<DocxReviewExporter>();
        services.AddSingleton<PdfReviewExporter>();
        services.AddSingleton<TicketExporter>();
        services.AddSingleton<FlatXlsxExporter>();
        services.AddSingleton<CombinedReportHtmlExporter>();
        services.AddSingleton<HostVulnerabilityReportExporter>();
        return services;
    }
}

public sealed record SessionExportOptions(
    bool IncludeHostVulnerabilityCounts = false,
    IReadOnlyList<Core.Models.HostVulnerabilitySummary>? HostCounts = null,
    bool UseStableFilenames = false);

public sealed class ExportOrchestrator(
    DocxReviewExporter docx,
    PdfReviewExporter pdf,
    FlatXlsxExporter xlsx,
    HostVulnerabilityReportExporter hostCountsExporter)
{
    public ExportResult ExportAll(
        VScanMagic.Review.Models.ReviewSession session,
        ReportOutputLayout layout,
        SessionExportOptions? options = null)
    {
        options ??= new SessionExportOptions();
        var topNDir = ReportPathResolver.GetTopNReportDirectory(layout);
        var supplementalDir = ReportPathResolver.GetSupplementalExportDirectory(layout);

        var useStable = options.UseStableFilenames || session.UseStableExportNames;
        var stamp = useStable ? null : ReportPathResolver.GetReportTimestamp();
        var companyName = string.IsNullOrWhiteSpace(session.ClientName) ? "Client" : session.ClientName.Trim();
        var reportTitle = ReviewExportLabels.GetReportTitle(session);
        var topLabel = ReviewExportLabels.GetTopNLabel(session);

        var docxPath = ReportPathResolver.GetSafeReportOutputPath(
            topNDir, companyName, ReportPathResolver.FormatExportSuffix($" {reportTitle}", stamp), "docx");
        var pdfPath = ReportPathResolver.GetSafeReportOutputPath(
            supplementalDir, companyName, ReportPathResolver.FormatExportSuffix($" {topLabel} Review (Client)", stamp), "pdf");
        var xlsxPath = ReportPathResolver.GetSafeReportOutputPath(
            supplementalDir, companyName, ReportPathResolver.FormatExportSuffix($" {topLabel} Data", stamp), "xlsx");

        docx.Export(session, docxPath);
        pdf.Export(session, pdfPath);
        xlsx.Export(session, xlsxPath);

        string? hostCountsPdf = null;
        string? hostCountsXlsx = null;
        if (options.IncludeHostVulnerabilityCounts && options.HostCounts is { Count: > 0 })
        {
            var hostRequest = new HostVulnerabilityReportRequest(
                session.ClientName,
                options.HostCounts,
                session.ScanDate,
                string.IsNullOrWhiteSpace(session.SourceFilePath) ? null : Path.GetFileName(session.SourceFilePath));
            var hostResult = hostCountsExporter.Export(hostRequest, supplementalDir, companyName, stamp, includePdf: true, includeXlsx: true);
            hostCountsPdf = hostResult.PdfPath;
            hostCountsXlsx = hostResult.XlsxPath;
        }

        return new ExportResult(
            layout.OutputDirectory,
            docxPath,
            pdfPath,
            xlsxPath,
            hostCountsPdf,
            hostCountsXlsx);
    }
}

public sealed record ExportResult(
    string OutputDirectory,
    string DocxPath,
    string PdfPath,
    string XlsxPath,
    string? HostCountsPdfPath = null,
    string? HostCountsXlsxPath = null);
