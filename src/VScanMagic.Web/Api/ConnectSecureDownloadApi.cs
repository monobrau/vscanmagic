using VScanMagic.ConnectSecure;
using VScanMagic.Core.Services;

namespace VScanMagic.Web.Api;

public static class ConnectSecureDownloadApi
{
    public static void MapConnectSecureDownloadApi(this WebApplication app)
    {
        app.MapPost("/api/connectsecure/download-standard", async (
            DownloadStandardRequest req,
            ConnectSecureReportService reportService,
            ConnectSecureClient client,
            SettingsService settings,
            CancellationToken ct) =>
        {
            var creds = settings.LoadConnectSecureCredentials();
            if (string.IsNullOrWhiteSpace(creds.BaseUrl))
                return Results.BadRequest("ConnectSecure credentials are not configured.");

            client.Configure(creds);

            if (string.IsNullOrWhiteSpace(req.CompanyId))
                return Results.BadRequest("CompanyId is required.");

            if (!int.TryParse(req.CompanyId, out var companyId))
                return Results.BadRequest("CompanyId must be numeric.");

            var userSettings = settings.LoadUserSettings();
            var outputDir = req.OutputDirectory
                ?? userSettings.LastOutputDirectory
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "VScanMagic", "Downloads");

            Directory.CreateDirectory(outputDir);

            var reports = BuildReportList(req);
            var progress = new List<string>();

            var result = await reportService.DownloadStandardReportsAsync(
                companyId,
                req.ClientName ?? "Client",
                req.ScanDate ?? DateTime.Now.ToString("yyyy-MM-dd"),
                outputDir,
                reports,
                new Progress<string>(m => progress.Add(m)),
                ct);

            userSettings.LastOutputDirectory = outputDir;
            settings.SaveUserSettings(userSettings);

            return Results.Ok(new
            {
                outputDirectory = outputDir,
                succeeded = result.Succeeded,
                failed = result.Failed,
                progress
            });
        });
    }

    private static List<StandardReportRequest> BuildReportList(DownloadStandardRequest req)
    {
        var all = StandardReportCatalog.DefaultCompanyReports;
        if (req.Reports is null || req.Reports.Count == 0)
            return all.ToList();

        var selected = req.Reports.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return all.Where(r => selected.Contains(r.Type)).ToList();
    }
}

public sealed class DownloadStandardRequest
{
    public string? CompanyId { get; set; }
    public string? ClientName { get; set; }
    public string? ScanDate { get; set; }
    public string? OutputDirectory { get; set; }
    public List<string>? Reports { get; set; }
}
