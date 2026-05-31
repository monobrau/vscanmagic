using VScanMagic.ConnectSecure;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Services;
using VScanMagic.Data;
using VScanMagic.Review;
using VScanMagic.Review.Models;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;

namespace VScanMagic.Web.Services;

public sealed class BulkReviewJobService
{
    private readonly IBulkReviewJobRepository _jobRepo;
    private readonly IReviewSessionRepository _sessionRepo;
    private readonly ConnectSecureReportService _reportService;
    private readonly VulnerabilityPipeline _pipeline;
    private readonly ReviewSessionFactory _sessionFactory;
    private readonly SettingsService _settings;
    private readonly ReportPathResolver _pathResolver;
    private readonly ReportFolderHistoryService _folderHistory;
    private readonly RmitPlusSettingsService _rmitPlusSettings;

    private readonly SemaphoreSlim _runLock = new(1, 1);
    private CancellationTokenSource? _runCts;
    private string? _activeJobId;

    public BulkReviewJobService(
        IBulkReviewJobRepository jobRepo,
        IReviewSessionRepository sessionRepo,
        ConnectSecureReportService reportService,
        VulnerabilityPipeline pipeline,
        ReviewSessionFactory sessionFactory,
        SettingsService settings,
        ReportPathResolver pathResolver,
        ReportFolderHistoryService folderHistory,
        RmitPlusSettingsService rmitPlusSettings)
    {
        _jobRepo = jobRepo;
        _sessionRepo = sessionRepo;
        _reportService = reportService;
        _pipeline = pipeline;
        _sessionFactory = sessionFactory;
        _settings = settings;
        _pathResolver = pathResolver;
        _folderHistory = folderHistory;
        _rmitPlusSettings = rmitPlusSettings;
    }

    public bool IsRunning => _activeJobId is not null;

    public string? ActiveJobId => _activeJobId;

    public Task<IReadOnlyList<BulkReviewJob>> ListJobsAsync(int limit = 20, CancellationToken ct = default) =>
        _jobRepo.ListAsync(limit, ct);

    public Task<BulkReviewJob?> GetJobAsync(string id, CancellationToken ct = default) =>
        _jobRepo.GetAsync(id, ct);

    public async Task<(bool Success, string Message, BulkReviewJob? Job)> StartJobAsync(
        BulkReviewJobRequest request,
        CancellationToken ct = default)
    {
        if (request.Companies.Count == 0)
            return (false, "Select at least one company.", null);

        if (string.IsNullOrWhiteSpace(request.ScanDate))
            return (false, "Scan date is required.", null);

        if (!await _runLock.WaitAsync(0, ct))
            return (false, "A bulk job is already running.", null);

        try
        {
            if (_activeJobId is not null)
                return (false, "A bulk job is already running.", null);

            var job = new BulkReviewJob
            {
                ScanDate = request.ScanDate.Trim(),
                Presenter = request.Presenter.Trim(),
                DefaultIsRmitPlus = request.DefaultIsRmitPlus,
                Status = BulkReviewJobStatus.Queued,
                Items = request.Companies.Select(c => new BulkReviewJobItem
                {
                    CompanyId = c.CompanyId,
                    CompanyName = c.CompanyName,
                    IsRmitPlus = c.IsRmitPlus
                }).ToList()
            };

            await _jobRepo.SaveAsync(job, ct);

            _runCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _activeJobId = job.Id;
            _ = Task.Run(async () =>
            {
                try
                {
                    await RunJobAsync(job.Id, _runCts.Token).ConfigureAwait(false);
                }
                finally
                {
                    _runLock.Release();
                }
            }, CancellationToken.None);

            return (true, "Bulk job started.", job);
        }
        catch
        {
            _runLock.Release();
            throw;
        }
    }

    public async Task CancelActiveJobAsync(CancellationToken ct = default)
    {
        _runCts?.Cancel();
        if (_activeJobId is null)
            return;

        var job = await _jobRepo.GetAsync(_activeJobId, ct);
        if (job is null || job.Status is BulkReviewJobStatus.Completed or BulkReviewJobStatus.Failed or BulkReviewJobStatus.Cancelled)
            return;

        job.Status = BulkReviewJobStatus.Cancelled;
        job.ErrorMessage = "Cancelled by user.";
        await _jobRepo.SaveAsync(job, ct);
    }

    private async Task RunJobAsync(string jobId, CancellationToken ct)
    {
        try
        {
            var job = await _jobRepo.GetAsync(jobId, ct);
            if (job is null)
                return;

            job.Status = BulkReviewJobStatus.Running;
            await _jobRepo.SaveAsync(job, ct);

            var userSettings = _settings.LoadUserSettings();
            var exportFilters = ReportFilters.FromUserSettings(userSettings);
            var exportTopN = exportFilters.TopN;
            var reports = StandardReportCatalog.DefaultCompanyReports.ToList();

            foreach (var item in job.Items)
            {
                if (ct.IsCancellationRequested)
                {
                    job.Status = BulkReviewJobStatus.Cancelled;
                    job.ErrorMessage = "Cancelled.";
                    await _jobRepo.SaveAsync(job, ct);
                    return;
                }

                if (item.Phase == BulkReviewItemPhase.Completed)
                    continue;

                try
                {
                    await ProcessItemAsync(job, item, userSettings, exportFilters, exportTopN, reports, ct);
                }
                catch (OperationCanceledException)
                {
                    job.Status = BulkReviewJobStatus.Cancelled;
                    job.ErrorMessage = "Cancelled.";
                    await _jobRepo.SaveAsync(job, ct);
                    return;
                }
                catch (Exception ex)
                {
                    item.Phase = BulkReviewItemPhase.Failed;
                    item.ErrorMessage = ex.Message;
                    item.StatusMessage = "Failed.";
                    await _jobRepo.SaveAsync(job, ct);
                }
            }

            job.Status = job.Items.Any(i => i.Phase == BulkReviewItemPhase.Completed)
                ? BulkReviewJobStatus.Completed
                : BulkReviewJobStatus.Failed;

            if (job.FailedCount > 0 && job.CompletedCount > 0)
                job.ErrorMessage = $"{job.CompletedCount} succeeded, {job.FailedCount} failed.";
            else if (job.FailedCount > 0 && job.CompletedCount == 0)
                job.ErrorMessage = "All clients failed.";

            await _jobRepo.SaveAsync(job, ct);
        }
        finally
        {
            _activeJobId = null;
            _runCts?.Dispose();
            _runCts = null;
        }
    }

    private async Task ProcessItemAsync(
        BulkReviewJob job,
        BulkReviewJobItem item,
        UserSettings userSettings,
        ReportFilters exportFilters,
        int exportTopN,
        IReadOnlyList<StandardReportRequest> reports,
        CancellationToken ct)
    {
        if (!int.TryParse(item.CompanyId, out var companyId))
            throw new InvalidOperationException($"Invalid company ID: {item.CompanyId}");

        var clientName = item.CompanyName.Trim();
        if (string.IsNullOrWhiteSpace(clientName))
            throw new InvalidOperationException("Company name is empty.");

        item.Phase = BulkReviewItemPhase.Downloading;
        item.StatusMessage = "Downloading reports...";
        item.ErrorMessage = null;
        await _jobRepo.SaveAsync(job, ct);

        var layout = _pathResolver.Resolve(
            userSettings,
            companyId,
            clientName,
            job.ScanDate,
            fallbackPath: userSettings.LastOutputDirectory);

        var downloadOptions = new ReportDownloadOptions(UseStableFilenames: false);
        string? allVulnsPath = null;
        const int maxAttempts = 2;
        Exception? lastCorruptError = null;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            var reportsForAttempt = attempt == 1 ? reports : StandardReportCatalog.AllVulnerabilitiesOnly;

            if (attempt > 1)
            {
                item.Phase = BulkReviewItemPhase.Downloading;
                item.StatusMessage = "Retrying All Vulnerabilities download…";
                await _jobRepo.SaveAsync(job, ct);
                TryDeleteFile(allVulnsPath);
            }

            var progress = new Progress<string>(msg =>
            {
                if (string.IsNullOrWhiteSpace(msg))
                    return;

                item.StatusMessage = msg;
                PersistJobProgress(job);
            });

            var downloadResult = await _reportService.DownloadStandardReportsAsync(
                companyId,
                clientName,
                layout,
                reportsForAttempt,
                downloadOptions,
                progress,
                ct).ConfigureAwait(false);

            var allVulns = downloadResult.Succeeded
                .FirstOrDefault(s => s.Name.Contains("All Vulnerabilities", StringComparison.OrdinalIgnoreCase));

            if (allVulns?.Path is null || !File.Exists(allVulns.Path))
            {
                var failDetail = downloadResult.Failed.Count > 0
                    ? string.Join("; ", downloadResult.Failed.Select(f => f.Error))
                    : "All Vulnerabilities report was not downloaded.";
                throw new InvalidOperationException(failDetail);
            }

            allVulnsPath = allVulns.Path;
            if (!IsLikelyValidXlsx(allVulnsPath))
            {
                if (attempt < maxAttempts)
                    continue;

                throw new InvalidOperationException(
                    "Downloaded All Vulnerabilities file is invalid or incomplete. " +
                    "Try Report downloader for this client, or run bulk prep again.");
            }

            item.OutputDirectory = layout.OutputDirectory;
            _folderHistory.Add(clientName, layout.OutputDirectory);
            ReportOutputDirectoryPersistence.UpdateLastOutputDirectory(userSettings, layout.OutputDirectory);
            _settings.SaveUserSettings(userSettings);

            item.Phase = BulkReviewItemPhase.Ingesting;
            item.StatusMessage = "Creating review session...";
            await _jobRepo.SaveAsync(job, ct);

            var ingestOptions = new VulnerabilityIngestOptions(
                SupplementDirectory: Path.GetDirectoryName(allVulnsPath),
                CompanyId: companyId);

            try
            {
                var result = await _pipeline.ProcessFileAsync(allVulnsPath, exportFilters, ingestOptions, ct).ConfigureAwait(false);
                if (result.Scored.AllFiltered.Count == 0)
                {
                    throw new InvalidOperationException(result.AllRecords.Count == 0
                        ? "No vulnerabilities found in All Vulnerabilities file."
                        : "No findings matched current severity filters.");
                }

                var session = _sessionFactory.CreateFromScoredResult(
                    clientName,
                    job.ScanDate,
                    result.Scored,
                    job.Presenter,
                    allVulnsPath,
                    item.CompanyId,
                    exportTopN,
                    item.IsRmitPlus);

                session.OutputDirectory = layout.OutputDirectory;
                session.UseStableExportNames = false;
                HostOsThresholdApplier.Apply(session, userSettings);
                DeliverableLinksResolver.ApplyDefaultsToSession(session, userSettings);
                _rmitPlusSettings.Set(clientName, item.IsRmitPlus);

                await _sessionRepo.SaveAsync(session, ct).ConfigureAwait(false);

                item.Phase = BulkReviewItemPhase.Completed;
                item.StatusMessage = "Review session created.";
                item.SessionId = session.Id;
                await _jobRepo.SaveAsync(job, ct);
                return;
            }
            catch (Exception ex) when (IsCorruptExcelError(ex))
            {
                lastCorruptError = ex;
                if (attempt < maxAttempts)
                    continue;
            }
        }

        throw new InvalidOperationException(
            "All Vulnerabilities XLSX could not be read (file may be truncated). " +
            "Use Report downloader for this client, then Start client review manually.",
            lastCorruptError);
    }

    private static bool IsLikelyValidXlsx(string path)
    {
        if (!File.Exists(path))
            return false;

        var info = new FileInfo(path);
        if (info.Length < 512)
            return false;

        Span<byte> header = stackalloc byte[2];
        using var stream = File.OpenRead(path);
        return stream.Read(header) == 2 && header[0] == 0x50 && header[1] == 0x4B;
    }

    private static bool IsCorruptExcelError(Exception ex)
    {
        for (var current = ex; current is not null; current = current.InnerException)
        {
            if (current.Message.Contains("corrupted data", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("invalid signature", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("central directory", StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrEmpty(path))
            return;

        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // Best effort before retry.
        }
    }

    private void PersistJobProgress(BulkReviewJob job)
    {
        try
        {
            _jobRepo.SaveAsync(job, CancellationToken.None).GetAwaiter().GetResult();
        }
        catch
        {
            // Progress updates are best-effort for UI polling.
        }
    }
}
