using System.Text.Json;
using VScanMagic.Core.IO;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureReportService(ConnectSecureClient client)
{
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(2);
    private const int MaxWaitSeconds = 600;

    public async Task<StandardReportDownloadResult> DownloadStandardReportsAsync(
        int companyId,
        string clientName,
        string scanDate,
        string outputDirectory,
        IEnumerable<StandardReportRequest>? reports = null,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        Directory.CreateDirectory(outputDirectory);
        var reportList = (reports ?? StandardReportCatalog.DefaultCompanyReports).ToList();
        var succeeded = new List<DownloadedReport>();
        var failed = new List<FailedReport>();
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");

        progress?.Report($"Loading standard report catalog for {clientName}...");
        var catalog = await client.GetStandardReportsAsync(companyId, ct);

        var pending = new List<(StandardReportRequest Report, string JobId, string Path)>();

        progress?.Report($"Creating {reportList.Count} report jobs...");
        foreach (var report in reportList)
        {
            var path = BuildOutputPath(outputDirectory, clientName, report, timestamp);
            try
            {
                var reportId = client.ResolveStandardReportId(report.Type, report.Extension, catalog, companyId);
                if (string.IsNullOrWhiteSpace(reportId))
                    throw new InvalidOperationException($"No standard report match for {report.Type} ({report.Extension}).");

                var displayName = StandardReportCatalog.DisplayNames.GetValueOrDefault(report.Type, report.Name);
                var jobId = await client.CreateReportJobAsync(
                    reportId, companyId, report.Extension, displayName, clientName, ct);
                pending.Add((report, jobId, path));
            }
            catch (Exception ex)
            {
                failed.Add(new FailedReport(report.Type, report.Name, ex.Message));
            }
        }

        if (pending.Count > 0)
        {
            await Task.Delay(TimeSpan.FromSeconds(5), ct);
            var start = DateTimeOffset.UtcNow;
            var isGlobal = companyId == 0;

            while (pending.Count > 0)
            {
                ct.ThrowIfCancellationRequested();
                if ((DateTimeOffset.UtcNow - start).TotalSeconds >= MaxWaitSeconds)
                {
                    foreach (var p in pending)
                        failed.Add(new FailedReport(p.Report.Type, p.Report.Name, "Timed out waiting for report generation."));
                    break;
                }

                progress?.Report($"Waiting for reports... ({pending.Count} pending)");
                var pollTasks = pending.Select(async p =>
                {
                    try
                    {
                        var url = await client.GetReportDownloadLinkAsync(p.JobId, isGlobal, companyId, ct);
                        if (string.IsNullOrWhiteSpace(url))
                            return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)null);

                        await client.DownloadFileFromUrlAsync(url, p.Path, ct);
                        ReportArchiveHelper.NormalizeDownloadedReportFile(p.Path);
                        return (Pending: p, Success: (DownloadedReport?)new DownloadedReport(p.Report.Type, p.Report.Name, p.Path), Failed: (FailedReport?)null);
                    }
                    catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)null);
                    }
                    catch (Exception ex)
                    {
                        return (Pending: p, Success: (DownloadedReport?)null, Failed: (FailedReport?)new FailedReport(p.Report.Type, p.Report.Name, ex.Message));
                    }
                }).ToList();

                var pollResults = await Task.WhenAll(pollTasks);
                var stillPending = new List<(StandardReportRequest Report, string JobId, string Path)>();
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
                    await Task.Delay(delay, ct);
                }
            }
        }

        return new StandardReportDownloadResult(succeeded, failed);
    }

    private static string BuildOutputPath(string dir, string clientName, StandardReportRequest report, string timestamp)
    {
        var safeClient = SanitizeFileName(clientName);
        var safeReport = SanitizeFileName(report.Name);
        return Path.Combine(dir, $"{safeClient} - {safeReport}_{timestamp}.{report.Extension}");
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? "Client" : name.Trim();
    }
}
