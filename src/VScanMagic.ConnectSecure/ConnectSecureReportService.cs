using VScanMagic.Core.IO;

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
        DownloadStandardReportsAsync(companyId, clientName, outputDirectory, reports, progress, ct);

    public async Task<StandardReportDownloadResult> DownloadStandardReportsAsync(
        int companyId,
        string clientName,
        string outputDirectory,
        IEnumerable<StandardReportRequest>? reports = null,
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
                items.Add(new CatalogReportDownloadRequest(reportId, displayName, report.Extension));
            }
            catch (Exception ex)
            {
                resolveFailed.Add(new FailedReport(report.Type, report.Name, ex.Message));
            }
        }

        if (items.Count == 0)
            return new StandardReportDownloadResult([], resolveFailed);

        var result = await DownloadCatalogReportsAsync(
            companyId, clientName, outputDirectory, items, progress, ct).ConfigureAwait(false);

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
        DownloadCatalogReportsAsync(companyId, clientName, outputDirectory, requests, progress, ct);

    public async Task<StandardReportDownloadResult> DownloadCatalogReportsAsync(
        int companyId,
        string clientName,
        string outputDirectory,
        IEnumerable<CatalogReportDownloadRequest> requests,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        Directory.CreateDirectory(outputDirectory);
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
            var path = BuildOutputPath(outputDirectory, clientName, report.Name, report.Extension, timestamp);
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

                progress?.Report($"Waiting for reports... ({pending.Count} pending)");
                var pollTasks = pending.Select(async p =>
                {
                    try
                    {
                        var url = await client.GetReportDownloadLinkAsync(p.JobId, isGlobal, companyId, ct).ConfigureAwait(false);
                        if (string.IsNullOrWhiteSpace(url))
                            return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)null);

                        await client.DownloadFileFromUrlAsync(url, p.Path, ct).ConfigureAwait(false);
                        ReportArchiveHelper.NormalizeDownloadedReportFile(p.Path);
                        return (Pending: p, Success: (DownloadedReport?)new DownloadedReport(
                            p.Report.ReportId, p.Report.Name, p.Path), Failed: (FailedReport?)null);
                    }
                    catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)null);
                    }
                    catch (Exception ex)
                    {
                        return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)new FailedReport(
                            p.Report.ReportId, p.Report.Name, ex.Message));
                    }
                }).ToList();

                var pollResults = await Task.WhenAll(pollTasks).ConfigureAwait(false);
                var stillPending = new List<(CatalogReportDownloadRequest Report, string JobId, string Path)>();
                foreach (var result in pollResults)
                {
                    if (result.Success is not null)
                        succeeded.Add(result.Success);
                    else if (result.Failed is not null)
                        failed.Add(result.Failed);
                    else
                        stillPending.Add(result.Pending);
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

    private static string BuildOutputPath(string dir, string clientName, string reportName, string extension, string timestamp)
    {
        var safeClient = SanitizeFileName(clientName);
        var safeReport = SanitizeFileName(reportName);
        return Path.Combine(dir, $"{safeClient} - {safeReport}_{timestamp}.{extension}");
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? "Client" : name.Trim();
    }
}
