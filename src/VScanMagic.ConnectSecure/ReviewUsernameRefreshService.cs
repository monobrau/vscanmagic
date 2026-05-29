using VScanMagic.Review;
using VScanMagic.Review.Models;

namespace VScanMagic.ConnectSecure;

public enum ReviewUsernameRefreshStatus
{
    Skipped,
    NoEmptyUsernames,
    NoCompanyId,
    NotConfigured,
    Refreshed
}

public sealed record ReviewUsernameRefreshResult(
    ReviewUsernameRefreshStatus Status,
    int HostnamesRequested = 0,
    int LookupMatches = 0,
    int UsernamesFilled = 0,
    string? Error = null);

public sealed class ReviewUsernameRefreshService(ConnectSecureUsernameLookupService lookup)
{
    /// <param name="exportFindingsOnly">
    /// When true, only hosts on findings in the export set are refreshed (Top N report and deliverables).
    /// </param>
    public async Task<ReviewUsernameRefreshResult> RefreshSessionAsync(
        ReviewSession session,
        int companyId,
        bool exportFindingsOnly = false,
        CancellationToken ct = default)
    {
        var hostnames = CollectHostnamesNeedingUsernames(session, exportFindingsOnly);
        if (hostnames.Count == 0)
            return new ReviewUsernameRefreshResult(ReviewUsernameRefreshStatus.NoEmptyUsernames);

        if (companyId <= 0)
            return new ReviewUsernameRefreshResult(ReviewUsernameRefreshStatus.NoCompanyId, hostnames.Count);

        if (!lookup.IsConfigured)
            return new ReviewUsernameRefreshResult(ReviewUsernameRefreshStatus.NotConfigured, hostnames.Count);

        try
        {
            var lookupResult = await lookup.GetUsernamesByHostnameAsync(companyId, hostnames, ct);
            var filled = ReviewUsernameLookup.ApplyEmptyUsernames(session, lookupResult);
            var matched = lookupResult.Values.Count(v => !string.IsNullOrWhiteSpace(v));
            return new ReviewUsernameRefreshResult(
                ReviewUsernameRefreshStatus.Refreshed,
                hostnames.Count,
                matched,
                filled);
        }
        catch (Exception ex)
        {
            return new ReviewUsernameRefreshResult(
                ReviewUsernameRefreshStatus.Skipped,
                hostnames.Count,
                Error: ex.Message);
        }
    }

    public static IReadOnlyList<string> CollectHostnamesNeedingUsernames(
        ReviewSession session,
        bool exportFindingsOnly = false)
    {
        var findings = exportFindingsOnly
            ? ReviewSessionRanker.GetExportFindings(session)
            : session.Findings;

        return findings
            .SelectMany(f => f.AffectedSystems ?? [])
            .Where(s => string.IsNullOrWhiteSpace(s.Username) && !string.IsNullOrWhiteSpace(s.HostName))
            .Select(s => s.HostName.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
