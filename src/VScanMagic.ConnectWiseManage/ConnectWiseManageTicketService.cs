using VScanMagic.Core.Paths;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Reports;
using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.ConnectWiseManage;

public sealed record ManageTicketCreateResult(
    bool Success,
    string Message,
    int? ManageTicketId = null,
    string? ManageTicketNumber = null,
    string? ManageTicketStatus = null);

public sealed record ManageTicketBatchResult(
    int Created,
    int Skipped,
    int Failed,
    IReadOnlyList<ManageTicketCreateResult> Results);

public sealed class ConnectWiseManageTicketService
{
    private static readonly TimeSpan CreateDelay = TimeSpan.FromMilliseconds(350);

    private readonly ConnectWiseManageClient _client;
    private readonly ConnectWiseCompanyMapService _companyMap;
    private readonly ConnectWiseManageSettingsStore _manageSettings;
    private readonly SettingsService _settings;
    private readonly RemediationRuleService _remediationRules;
    private readonly ReportPathResolver _pathResolver;

    public ConnectWiseManageTicketService(
        ConnectWiseManageClient client,
        ConnectWiseCompanyMapService companyMap,
        ConnectWiseManageSettingsStore manageSettings,
        SettingsService settings,
        RemediationRuleService remediationRules,
        ReportPathResolver pathResolver)
    {
        _client = client;
        _companyMap = companyMap;
        _manageSettings = manageSettings;
        _settings = settings;
        _remediationRules = remediationRules;
        _pathResolver = pathResolver;
    }

    public async Task<ManageTicketBatchResult> CreateTicketsForSessionAsync(
        ReviewSession session,
        CancellationToken ct = default)
    {
        EnsureConfigured(session);

        var options = _manageSettings.LoadOptions();
        if (options.DefaultBoardId <= 0 || options.DefaultStatusId <= 0)
            throw new InvalidOperationException("Set default Manage board and status on Settings.");

        if (!_companyMap.TryGetManageCompanyId(session.CompanyId ?? "", out var manageCompanyId, out _))
            throw new InvalidOperationException(
                $"No ConnectWise Manage company mapping for ConnectSecure company {session.CompanyId}. Add it on Client mappings.");

        var layout = ResolveLayout(session);
        var sections = TicketInstructionBuilder.BuildSections(session, _remediationRules, layout.ReportsPathPartial);
        var sectionByRank = sections.ToDictionary(s => s.Number, s => s);
        var findings = ReviewSessionRanker.GetExportFindings(session);

        var results = new List<ManageTicketCreateResult>();
        var created = 0;
        var skipped = 0;
        var failed = 0;

        for (var i = 0; i < findings.Count; i++)
        {
            var finding = findings[i];
            var sectionNumber = i + 1;

            if (!TimeEstimateModifierHelper.IsTicketGenerated(
                    finding.AfterHours, finding.TicketGenerated, finding.ThirdParty))
            {
                skipped++;
                results.Add(new ManageTicketCreateResult(false, "Not flagged for ticket creation."));
                continue;
            }

            if (finding.ManageTicketId is not null)
            {
                skipped++;
                results.Add(new ManageTicketCreateResult(
                    true,
                    $"Already created (#{finding.ManageTicketNumber ?? finding.ManageTicketId.ToString()}).",
                    finding.ManageTicketId,
                    finding.ManageTicketNumber,
                    finding.ManageTicketStatus));
                continue;
            }

            if (!sectionByRank.TryGetValue(sectionNumber, out var section))
            {
                failed++;
                results.Add(new ManageTicketCreateResult(false, "Ticket instruction section not found."));
                continue;
            }

            try
            {
                var ticket = await _client.CreateTicketAsync(new ManageTicketCreateRequest
                {
                    Summary = section.Subject,
                    InitialDescription = section.BodyText,
                    Board = new ManageReference { Id = options.DefaultBoardId },
                    Status = new ManageReference { Id = options.DefaultStatusId },
                    Company = new ManageReference { Id = manageCompanyId }
                }, ct).ConfigureAwait(false);

                finding.ManageTicketId = ticket.Id;
                finding.ManageTicketNumber = ticket.Id.ToString();
                finding.ManageTicketStatus = ticket.Status?.Name ?? "New";
                finding.ManageTicketCreatedAt = DateTimeOffset.Now;
                finding.TicketGenerated = true;

                created++;
                results.Add(new ManageTicketCreateResult(
                    true,
                    $"Created ticket #{finding.ManageTicketNumber}.",
                    finding.ManageTicketId,
                    finding.ManageTicketNumber,
                    finding.ManageTicketStatus));

                await Task.Delay(CreateDelay, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                failed++;
                results.Add(new ManageTicketCreateResult(false, ex.Message));
            }
        }

        return new ManageTicketBatchResult(created, skipped, failed, results);
    }

    public async Task<int> RefreshTicketsForSessionAsync(ReviewSession session, CancellationToken ct = default)
    {
        EnsureConfigured(session);

        var findings = ReviewSessionRanker.GetExportFindings(session);
        var refreshed = 0;

        foreach (var finding in findings.Where(f => f.ManageTicketId is not null))
        {
            try
            {
                var ticket = await _client.GetTicketAsync(finding.ManageTicketId!.Value, ct).ConfigureAwait(false);
                finding.ManageTicketNumber = ticket.Id.ToString();
                finding.ManageTicketStatus = ticket.Status?.Name ?? finding.ManageTicketStatus;
                refreshed++;
                await Task.Delay(CreateDelay, ct).ConfigureAwait(false);
            }
            catch
            {
                // Keep existing values when refresh fails for one ticket.
            }
        }

        return refreshed;
    }

    private void EnsureConfigured(ReviewSession session)
    {
        if (!_client.IsConfigured)
            throw new InvalidOperationException("ConnectWise Manage is not configured on Settings.");

        if (string.IsNullOrWhiteSpace(session.CompanyId))
            throw new InvalidOperationException("Review session has no ConnectSecure company ID.");
    }

    private ReportOutputLayout ResolveLayout(ReviewSession session)
    {
        var userSettings = _settings.LoadUserSettings();
        var companyId = int.TryParse(session.CompanyId, out var id) ? id : 0;
        return SessionOutputLayoutResolver.ResolveForSession(
            _pathResolver,
            userSettings,
            session.ClientName,
            session.ScanDate,
            companyId,
            session.OutputDirectory,
            session.SourceFilePath,
            fallbackPath: userSettings.LastOutputDirectory);
    }
}
