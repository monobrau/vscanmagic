using VScanMagic.Core.Models;
using VScanMagic.Core.Nvd;
using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review.Services;

public sealed class CveEnrichmentService(NvdCveLookupService nvdLookup)
{
    public async Task<bool> EnrichFindingAsync(
        ReviewFinding finding,
        UserSettings settings,
        bool applyToRemediation,
        CancellationToken ct = default)
    {
        var cveIds = CveEnrichmentPolicy.GetCveIds(finding);
        if (cveIds.Count == 0)
            return false;

        var summary = await nvdLookup.GetRemediationSummaryForListAsync(cveIds, settings.NvdApiKey, ct)
            .ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(summary))
            return false;

        finding.NvdEnrichment = summary;
        if (applyToRemediation)
            ApplyToRemediation(finding, summary);

        return true;
    }

    public async Task<int> EnrichOnIngestAsync(
        ReviewSession session,
        UserSettings settings,
        CancellationToken ct = default)
    {
        if (!settings.NvdEnrichOnIngest)
            return 0;

        return await EnrichFindingsAsync(
            session.Findings.Where(f => f.IncludeInExport),
            settings,
            ct).ConfigureAwait(false);
    }

    /// <summary>Enrich export-set findings that still need NVD (e.g. on review load).</summary>
    public async Task<int> EnrichMissingExportAsync(
        ReviewSession session,
        UserSettings settings,
        CancellationToken ct = default)
    {
        return await EnrichFindingsAsync(
            session.Findings.Where(f => f.IncludeInExport && CveEnrichmentPolicy.ShouldEnrich(f)),
            settings,
            ct).ConfigureAwait(false);
    }

    public async Task<int> EnrichMissingAsync(
        ReviewSession session,
        UserSettings settings,
        CancellationToken ct = default)
    {
        return await EnrichFindingsAsync(
            session.Findings.Where(CveEnrichmentPolicy.ShouldEnrich),
            settings,
            ct).ConfigureAwait(false);
    }

    private async Task<int> EnrichFindingsAsync(
        IEnumerable<ReviewFinding> findings,
        UserSettings settings,
        CancellationToken ct)
    {
        var enriched = 0;
        foreach (var finding in findings)
        {
            if (CveEnrichmentPolicy.GetCveIds(finding).Count == 0)
                continue;

            if (CveEnrichmentPolicy.HasActionableFix(finding) &&
                !CveReferenceHelper.IsCveOnlyProduct(finding.Product))
                continue;

            if (await EnrichFindingAsync(finding, settings, applyToRemediation: true, ct).ConfigureAwait(false))
                enriched++;
        }

        return enriched;
    }

    private static void ApplyToRemediation(ReviewFinding finding, string summary)
    {
        if (CveEnrichmentPolicy.HasActionableFix(finding))
            return;

        if (FindingRemediationExport.IsRemediationEdited(finding))
            return;

        var revised = CveEnrichmentPolicy.AppendNvdContext(FindingExportDetails.GetRemediationText(finding), finding);
        finding.RevisedRemediation = revised;
        if (string.IsNullOrWhiteSpace(finding.OriginalRemediation))
            finding.OriginalRemediation = revised;
    }
}
