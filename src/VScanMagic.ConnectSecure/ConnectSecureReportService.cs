using VScanMagic.Core.IO;
using VScanMagic.Core.Paths;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureReportService(ConnectSecureClient client)
{
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(2);
    private const int MaxWaitSeconds = 600;

    public Task<StandardReportDownloadResult> DownloadStandardReportsAsync(
        int companyId,
        string clientName,
        string scanDate,
        string outputDirectory,
        IEnumerable<StandardReportRequest>? reports = null,
        IProgress<string>? progress = null,
        CancellationToken ct = default) =>
        DownloadStandardReportsAsync(companyId, clientName, FlatLayout(outputDirectory), reports, progress, ct);

    public async Task<StandardReportDownloadResult> DownloadStandardReportsAsync(
        int companyId,
        string clientName,
        ReportOutputLayout layout,
        IEnumerable<StandardReportRequest>? reports = null,
        IProgress<string>? progress = null,
        CancellationToken ct = default) =>
        await DownloadStandardReportsAsync(
            companyId, clientName, layout, reports, new ReportDownloadOptions(), progress, ct).ConfigureAwait(false);

    public async Task<StandardReportDownloadResult> DownloadStandardReportsAsync(
        int companyId,
        string clientName,
        ReportOutputLayout layout,
        IEnumerable<StandardReportRequest>? reports,
        ReportDownloadOptions downloadOptions,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var reportList = (reports ?? StandardReportCatalog.DefaultCompanyReports).ToList();
        progress?.Report($"Loading standard report catalog for {clientName}...");
        var catalog = await client.GetStandardReportsAsync(companyId, ct: ct).ConfigureAwait(false);

        var items = new List<CatalogReportDownloadRequest>();
        var resolveFailed = new List<FailedReport>();
        foreach (var report in reportList)
        {
            try
            {
                var reportId = client.ResolveStandardReportId(report.Type, report.Extension, catalog, companyId);
                if (string.IsNullOrWhiteSpace(reportId))
                    throw new InvalidOperationException($"No standard report match for {report.Type} ({report.Extension}).");

                var displayName = StandardReportCatalog.DisplayNames.GetValueOrDefault(report.Type, report.Name);
                items.Add(new CatalogReportDownloadRequest(reportId, displayName, report.Extension, report.Type));
            }
            catch (Exception ex)
            {
                resolveFailed.Add(new FailedReport(report.Type, report.Name, ex.Message));
            }
        }

        if (items.Count == 0)
            return new StandardReportDownloadResult([], resolveFailed);

        var result = await DownloadCatalogReportsAsync(
            companyId, clientName, layout, items, downloadOptions, progress, ct).ConfigureAwait(false);

        if (resolveFailed.Count == 0)
            return result;

        return new StandardReportDownloadResult(
            result.Succeeded,
            resolveFailed.Concat(result.Failed).ToList());
    }

    public Task<StandardReportDownloadResult> DownloadCatalogReportsAsync(
        int companyId,
        string clientName,
        string scanDate,
        string outputDirectory,
        IEnumerable<CatalogReportDownloadRequest> requests,
        IProgress<string>? progress = null,
        CancellationToken ct = default) =>
        DownloadCatalogReportsAsync(companyId, clientName, FlatLayout(outputDirectory), requests, progress, ct);

    public async Task<StandardReportDownloadResult> DownloadCatalogReportsAsync(
        int companyId,
        string clientName,
        ReportOutputLayout layout,
        IEnumerable<CatalogReportDownloadRequest> requests,
        IProgress<string>? progress = null,
        CancellationToken ct = default) =>
        await DownloadCatalogReportsAsync(
            companyId, clientName, layout, requests, new ReportDownloadOptions(), progress, ct).ConfigureAwait(false);

    public async Task<StandardReportDownloadResult> DownloadCatalogReportsAsync(
        int companyId,
        string clientName,
        ReportOutputLayout layout,
        IEnumerable<CatalogReportDownloadRequest> requests,
        ReportDownloadOptions downloadOptions,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var reportList = requests.ToList();
        if (reportList.Count == 0)
            return new StandardReportDownloadResult([], []);

        var succeeded = new List<DownloadedReport>();
        var failed = new List<FailedReport>();
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");

        progress?.Report($"Creating {reportList.Count} report jobs...");
        var pending = new List<(CatalogReportDownloadRequest Report, string JobId, string Path)>();

        foreach (var report in reportList)
        {
            var downloadDir = ReportPathResolver.GetDownloadDirectory(layout, report.ReportType);
            var path = BuildOutputPath(
                downloadDir,
                clientName,
                report.Name,
                report.Extension,
                downloadOptions.UseStableFilenames ? null : timestamp);
            try
            {
                var jobId = await client.CreateReportJobAsync(
                    report.ReportId, companyId, report.Extension, report.Name, clientName, ct).ConfigureAwait(false);
                pending.Add((report, jobId, path));
            }
            catch (Exception ex)
            {
                failed.Add(new FailedReport(report.ReportId, report.Name, ex.Message));
            }
        }

        if (pending.Count > 0)
        {
            await Task.Delay(TimeSpan.FromSeconds(5), ct).ConfigureAwait(false);
            var start = DateTimeOffset.UtcNow;
            var isGlobal = companyId == 0;

            while (pending.Count > 0)
            {
                ct.ThrowIfCancellationRequested();
                if ((DateTimeOffset.UtcNow - start).TotalSeconds >= MaxWaitSeconds)
                {
                    foreach (var p in pending)
                        failed.Add(new FailedReport(p.Report.ReportId, p.Report.Name, "Timed out waiting for report generation."));
                    break;
                }

                var elapsed = (int)(DateTimeOffset.UtcNow - start).TotalSeconds;
                var pendingNames = string.Join(", ", pending.Select(p => p.Report.Name));
                progress?.Report(
                    pending.Count == 1
                        ? $"Waiting for {pending[0].Report.Name}… ({elapsed}s — ConnectSecure is still generating)"
                        : $"Waiting for reports… ({elapsed}s, {pending.Count} pending: {pendingNames})");

                var stillPending = new List<(CatalogReportDownloadRequest Report, string JobId, string Path)>();
                foreach (var p in pending)
                {
                    try
                    {
                        var url = await client.GetReportDownloadLinkAsync(p.JobId, isGlobal, companyId, ct).ConfigureAwait(false);
                        if (string.IsNullOrWhiteSpace(url))
                        {
                            stillPending.Add(p);
                            continue;
                        }

                        progress?.Report($"Downloading {p.Report.Name}…");
                        await client.DownloadFileFromUrlAsync(url, p.Path, ct).ConfigureAwait(false);
                        ReportArchiveHelper.NormalizeDownloadedReportFile(p.Path);
                        succeeded.Add(new DownloadedReport(p.Report.ReportId, p.Report.Name, p.Path));
                    }
                    catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        stillPending.Add(p);
                    }
                    catch (Exception ex)
                    {
                        failed.Add(new FailedReport(p.Report.ReportId, p.Report.Name, ex.Message));
                    }
                }

                pending = stillPending;
                if (pending.Count > 0)
                {
                    var elapsedSeconds = (DateTimeOffset.UtcNow - start).TotalSeconds;
                    var delay = elapsedSeconds >= 30 ? TimeSpan.FromSeconds(4) : PollInterval;
                    await Task.Delay(delay, ct).ConfigureAwait(false);
                }
            }
        }

        return new StandardReportDownloadResult(succeeded, failed);
    }

    private static string BuildOutputPath(
        string dir,
        string clientName,
        string reportName,
        string extension,
        string? timestamp)
    {
        var safeClient = SanitizeFileName(clientName);
        var safeReport = SanitizeFileName(reportName);
        var suffix = string.IsNullOrWhiteSpace(timestamp) ? "" : $"_{timestamp}";
        return Path.Combine(dir, $"{safeClient} - {safeReport}{suffix}.{extension}");
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? "Client" : name.Trim();
    }

    private static ReportOutputLayout FlatLayout(string outputDirectory)
    {
        var dir = Path.GetFullPath(outputDirectory.Trim());
        Directory.CreateDirectory(dir);
        return new ReportOutputLayout
        {
            OutputDirectory = dir,
            TextOutputDirectory = dir,
            UsesStructuredPaths = false,
            UsesMiscSubfolder = false
        };
    }
}
