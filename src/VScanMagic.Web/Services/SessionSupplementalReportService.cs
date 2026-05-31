using VScanMagic.ConnectSecure;
using VScanMagic.Core.Paths;
using VScanMagic.Review.Models;

namespace VScanMagic.Web.Services;

public sealed class SessionSupplementalReportService(ConnectSecureReportService reportService)
{
    public async Task<StandardReportDownloadResult> EnsureSupplementalReportsAsync(
        ReviewSession session,
        ReportOutputLayout layout,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        if (!int.TryParse(session.CompanyId, out var companyId) || companyId <= 0)
        {
            return new StandardReportDownloadResult(
                [],
                [new FailedReport("", "ConnectSecure", "Session has no ConnectSecure company ID — skipped supplemental report download.")]);
        }

        var clientName = session.ClientName.Trim();
        if (string.IsNullOrWhiteSpace(clientName))
        {
            return new StandardReportDownloadResult(
                [],
                [new FailedReport("", "ConnectSecure", "Session has no client name — skipped supplemental report download.")]);
        }

        var reports = ConnectSecureReportService.FilterMissingReports(
            layout,
            clientName,
            StandardReportCatalog.SupplementalCompanyReports,
            session.UseStableExportNames);

        if (reports.Count == 0)
        {
            progress?.Report("ConnectSecure supplemental reports already on disk.");
            return new StandardReportDownloadResult([], []);
        }

        progress?.Report($"Downloading {reports.Count} ConnectSecure supplemental report(s)…");
        return await reportService.DownloadStandardReportsAsync(
            companyId,
            clientName,
            layout,
            reports,
            new ReportDownloadOptions(UseStableFilenames: session.UseStableExportNames),
            progress,
            ct).ConfigureAwait(false);
    }
}
