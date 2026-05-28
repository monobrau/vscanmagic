using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;
using VScanMagic.Data;
using VScanMagic.Reports;
using VScanMagic.Review.Models;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;

namespace VScanMagic.Web.Api;

public static class ReviewSessionApi
{
    public static void MapReviewSessionApi(this WebApplication app)
    {
        var group = app.MapGroup("/api/review-sessions");

        group.MapGet("/", async (IReviewSessionRepository repo, bool includeArchived, CancellationToken ct) =>
            Results.Ok(await repo.ListAsync(includeArchived, ct)));

        group.MapDelete("/{id}", async (string id, IReviewSessionRepository repo, CancellationToken ct) =>
        {
            var existing = await repo.GetAsync(id, ct);
            if (existing is null) return Results.NotFound();
            await repo.DeleteAsync(id, ct);
            return Results.NoContent();
        });

        group.MapGet("/{id}", async (string id, IReviewSessionRepository repo, CancellationToken ct) =>
        {
            var session = await repo.GetAsync(id, ct);
            return session is null ? Results.NotFound() : Results.Ok(session);
        });

        group.MapPost("/", async (
            CreateSessionRequest request,
            VulnerabilityPipeline pipeline,
            ReviewSessionFactory factory,
            IReviewSessionRepository repo,
            SettingsService settings,
            CancellationToken ct) =>
        {
            if (string.IsNullOrWhiteSpace(request.ExcelPath) || !File.Exists(request.ExcelPath))
                return Results.BadRequest("ExcelPath is required and must exist.");

            var userSettings = settings.LoadUserSettings();
            var filters = request.Filters ?? ReportFilters.FromUserSettings(userSettings);
            var result = pipeline.ProcessFile(request.ExcelPath, filters);
            var session = factory.CreateFromScoredResult(
                request.ClientName,
                request.ScanDate,
                result.Scored,
                request.Presenter ?? userSettings.PreparedBy,
                request.ExcelPath,
                request.CompanyId,
                filters.TopN);
            await repo.SaveAsync(session, ct);
            return Results.Created($"/api/review-sessions/{session.Id}", session);
        });

        group.MapPatch("/{id}", async (string id, ReviewSession patch, IReviewSessionRepository repo, CancellationToken ct) =>
        {
            var existing = await repo.GetAsync(id, ct);
            if (existing is null) return Results.NotFound();
            existing.Findings = patch.Findings;
            existing.ClientName = patch.ClientName ?? existing.ClientName;
            existing.ScanDate = patch.ScanDate ?? existing.ScanDate;
            existing.Presenter = patch.Presenter ?? existing.Presenter;
            existing.IsRmitPlus = patch.IsRmitPlus;
            existing.ArchivedAt = patch.ArchivedAt;
            await repo.SaveAsync(existing, ct);
            return Results.Ok(existing);
        });

        group.MapPost("/{id}/export/docx", async (string id, ExportRequest req, IReviewSessionRepository repo, DocxReviewExporter exporter, CancellationToken ct) =>
        {
            var session = await repo.GetAsync(id, ct);
            if (session is null) return Results.NotFound();
            var path = req.OutputPath ?? Path.Combine(req.OutputDirectory ?? Path.GetTempPath(), $"review_{id}.docx");
            exporter.Export(session, path);
            return Results.Ok(new { path });
        });

        group.MapPost("/{id}/export/pdf", async (string id, ExportRequest req, IReviewSessionRepository repo, PdfReviewExporter exporter, CancellationToken ct) =>
        {
            var session = await repo.GetAsync(id, ct);
            if (session is null) return Results.NotFound();
            var path = req.OutputPath ?? Path.Combine(req.OutputDirectory ?? Path.GetTempPath(), $"review_{id}.pdf");
            exporter.Export(session, path);
            return Results.Ok(new { path });
        });

        group.MapPost("/{id}/export/all", async (
            string id,
            ExportRequest req,
            IReviewSessionRepository repo,
            ExportOrchestrator orchestrator,
            SettingsService settings,
            ReportPathResolver pathResolver,
            CancellationToken ct) =>
        {
            var session = await repo.GetAsync(id, ct);
            if (session is null) return Results.NotFound();

            var userSettings = settings.LoadUserSettings();
            var companyId = int.TryParse(session.CompanyId, out var parsedId) ? parsedId : 0;
            var layout = pathResolver.Resolve(
                userSettings,
                companyId,
                session.ClientName,
                session.ScanDate,
                req.OutputDirectory);

            var result = orchestrator.ExportAll(session, layout);
            return Results.Ok(result);
        });
    }
}

public sealed class CreateSessionRequest
{
    public string ClientName { get; set; } = "";
    public string ScanDate { get; set; } = "";
    public string ExcelPath { get; set; } = "";
    public string? Presenter { get; set; }
    public string? CompanyId { get; set; }
    public ReportFilters? Filters { get; set; }
}

public sealed class ExportRequest
{
    public string? OutputDirectory { get; set; }
    public string? OutputPath { get; set; }
}
