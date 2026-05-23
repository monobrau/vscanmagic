using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;
using VScanMagic.Data;
using VScanMagic.Reports;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;

namespace VScanMagic.Web.Api;

public static class LegacyReportApi
{
    public static void MapLegacyReportApi(this WebApplication app)
    {
        var apiKey = Environment.GetEnvironmentVariable("VSCANMAGIC_API_KEY");

        var group = app.MapGroup("/api/reports");

        group.MapPost("/executive-summary", async (
            LegacyReportRequest req,
            VulnerabilityPipeline pipeline,
            ReviewSessionFactory factory,
            ExportOrchestrator exporter,
            SettingsService settings,
            ReportPathResolver pathResolver,
            CancellationToken ct) =>
        {
            if (!Authorize(apiKey, req.ApiKey)) return Results.Unauthorized();
            return await GenerateLegacyAsync(req, pipeline, factory, exporter, settings, pathResolver, "docx", ct);
        });

        group.MapPost("/pending-epss", async (
            LegacyReportRequest req,
            VulnerabilityPipeline pipeline,
            FlatXlsxExporter xlsx,
            SettingsService settings,
            CancellationToken ct) =>
        {
            if (!Authorize(apiKey, req.ApiKey)) return Results.Unauthorized();
            var filters = req.Filters ?? ReportFilters.FromUserSettings(settings.LoadUserSettings());
            var result = pipeline.ProcessFile(req.InputPath, filters);
            var path = req.OutputPath ?? Path.Combine(Path.GetTempPath(), $"pending_epss_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            xlsx.ExportSummaryFromRecords(result.AllRecords, path);
            return Results.Ok(new { path });
        });

        group.MapPost("/all-vulnerabilities", async (
            LegacyReportRequest req,
            VulnerabilityPipeline pipeline,
            FlatXlsxExporter xlsx,
            SettingsService settings,
            CancellationToken ct) =>
        {
            if (!Authorize(apiKey, req.ApiKey)) return Results.Unauthorized();
            var filters = req.Filters ?? ReportFilters.FromUserSettings(settings.LoadUserSettings());
            var result = pipeline.ProcessFile(req.InputPath, filters);
            var path = req.OutputPath ?? Path.Combine(Path.GetTempPath(), $"all_vulns_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            xlsx.ExportSummaryFromRecords(result.AllRecords, path);
            return Results.Ok(new { path });
        });
    }

    private static bool Authorize(string? expected, string? provided)
    {
        if (string.IsNullOrWhiteSpace(expected)) return true;
        return string.Equals(expected, provided, StringComparison.Ordinal);
    }

    private static async Task<IResult> GenerateLegacyAsync(
        LegacyReportRequest req,
        VulnerabilityPipeline pipeline,
        ReviewSessionFactory factory,
        ExportOrchestrator exporter,
        SettingsService settings,
        ReportPathResolver pathResolver,
        string format,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.InputPath) || !File.Exists(req.InputPath))
            return Results.BadRequest("InputPath required.");

        var userSettings = settings.LoadUserSettings();
        var filters = req.Filters ?? ReportFilters.FromUserSettings(userSettings);
        var result = pipeline.ProcessFile(req.InputPath, filters);
        var session = factory.CreateFromTopVulnerabilities(
            req.ClientName ?? "Client",
            req.ScanDate ?? DateTime.Now.ToString("yyyy-MM-dd"),
            result.TopVulnerabilities,
            userSettings.PreparedBy,
            req.InputPath);

        var companyId = 0;
        var layout = pathResolver.Resolve(
            userSettings,
            companyId,
            session.ClientName,
            session.ScanDate,
            req.OutputDirectory);
        var export = exporter.ExportAll(session, layout);
        return Results.Ok(new { sessionId = session.Id, outputDirectory = export.OutputDirectory, docx = export.DocxPath, pdf = export.PdfPath });
    }
}

public sealed class LegacyReportRequest
{
    public string InputPath { get; set; } = "";
    public string? OutputPath { get; set; }
    public string? OutputDirectory { get; set; }
    public string? ClientName { get; set; }
    public string? ScanDate { get; set; }
    public string? ApiKey { get; set; }
    public ReportFilters? Filters { get; set; }
}
