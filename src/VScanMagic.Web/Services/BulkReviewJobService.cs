using VScanMagic.ConnectSecure;
using VScanMagic.Core.IO;
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
    private static readonly TimeSpan ProgressPersistInterval = TimeSpan.FromSeconds(2);
    private static readonly IReadOnlyList<StandardReportRequest> BulkPrepReports =
        StandardReportCatalog.AllVulnerabilitiesOnly;

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
    private readonly object _progressLock = new();
    private CancellationTokenSource? _runCts;
    private string? _activeJobId;
    private volatile bool _cancelRequested;
    private DateTimeOffset _lastProgressPersist = DateTimeOffset.MinValue;

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

    public Task RecoverInterruptedJobsAsync(CancellationToken ct = default) =>
        _jobRepo.RecoverInterruptedJobsAsync(ct);

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
            {
                _runLock.Release();
                return (false, "A bulk job is already running.", null);
            }

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

            _cancelRequested = false;
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
        _cancelRequested = true;
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
        var userSettingsDirty = false;
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

            foreach (var item in job.Items)
            {
                if (IsJobCancelled(ct))
                {
                    await MarkJobCancelledAsync(job, ct);
                    return;
                }

                if (item.Phase == BulkReviewItemPhase.Completed)
                    continue;

                try
                {
                    userSettingsDirty |= await ProcessItemAsync(job, item, userSettings, exportFilters, exportTopN, ct);
                }
                catch (OperationCanceledException)
                {
                    await MarkJobCancelledAsync(job, ct);
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

            if (userSettingsDirty)
                _settings.SaveUserSettings(userSettings);

            await FinalizeJobAsync(job, ct);
        }
        finally
        {
            _activeJobId = null;
            _cancelRequested = false;
            _runCts?.Dispose();
            _runCts = null;
        }
    }

    private async Task FinalizeJobAsync(BulkReviewJob job, CancellationToken ct)
    {
        if (IsJobCancelled(ct) || job.Status == BulkReviewJobStatus.Cancelled)
        {
            await MarkJobCancelledAsync(job, ct);
            return;
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

    private async Task MarkJobCancelledAsync(BulkReviewJob job, CancellationToken ct)
    {
        job.Status = BulkReviewJobStatus.Cancelled;
        job.ErrorMessage ??= "Cancelled.";
        await _jobRepo.SaveAsync(job, ct);
    }

    private bool IsJobCancelled(CancellationToken ct) =>
        _cancelRequested || ct.IsCancellationRequested;

    private async Task<bool> ProcessItemAsync(
        BulkReviewJob job,
        BulkReviewJobItem item,
        UserSettings userSettings,
        ReportFilters exportFilters,
        int exportTopN,
        CancellationToken ct)
    {
        if (!int.TryParse(item.CompanyId, out var companyId))
            throw new InvalidOperationException($"Invalid company ID: {item.CompanyId}");

        var clientName = item.CompanyName.Trim();
        if (string.IsNullOrWhiteSpace(clientName))
            throw new InvalidOperationException("Company name is empty.");

        item.Phase = BulkReviewItemPhase.Downloading;
        item.StatusMessage = "Downloading All Vulnerabilities report...";
        item.ErrorMessage = null;
        ResetProgressThrottle();
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
        var settingsDirty = false;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            if (attempt > 1)
            {
                item.Phase = BulkReviewItemPhase.Downloading;
                item.StatusMessage = "Retrying All Vulnerabilities download (ConnectSecure may still be preparing the file)…";
                ResetProgressThrottle();
                await _jobRepo.SaveAsync(job, ct);
                XlsxFileValidator.TryDeleteFile(allVulnsPath);
            }

            var progress = new Progress<string>(msg =>
            {
                if (string.IsNullOrWhiteSpace(msg))
                    return;

                item.StatusMessage = msg;
                PersistJobProgressThrottled(job);
            });

            var downloadResult = await _reportService.DownloadStandardReportsAsync(
                companyId,
                clientName,
                layout,
                BulkPrepReports,
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
            if (!XlsxFileValidator.IsLikelyValidXlsx(allVulnsPath))
            {
                if (attempt < maxAttempts)
                    continue;

                throw new InvalidOperationException(
                    "All Vulnerabilities file is not ready to read yet. ConnectSecure may still be generating it for large clients — try Report downloader or run bulk prep again.");
            }

            item.OutputDirectory = layout.OutputDirectory;
            _folderHistory.Add(clientName, layout.OutputDirectory);
            ReportOutputDirectoryPersistence.UpdateLastOutputDirectory(userSettings, layout.OutputDirectory);
            settingsDirty = true;

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
                return settingsDirty;
            }
            catch (Exception ex) when (XlsxFileValidator.IsCorruptExcelError(ex))
            {
                lastCorruptError = ex;
                if (attempt < maxAttempts)
                    continue;
            }
        }

        throw new InvalidOperationException(
            "All Vulnerabilities XLSX could not be read after waiting for ConnectSecure to finish. " +
            "Use Report downloader for this client (allow extra time for large reports), then start review manually.",
            lastCorruptError);
    }

    private void ResetProgressThrottle() =>
        _lastProgressPersist = DateTimeOffset.MinValue;

    private void PersistJobProgressThrottled(BulkReviewJob job)
    {
        var now = DateTimeOffset.UtcNow;
        lock (_progressLock)
        {
            if (now - _lastProgressPersist < ProgressPersistInterval)
                return;

            _lastProgressPersist = now;
        }

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
