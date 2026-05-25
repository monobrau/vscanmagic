using VScanMagic.Core.Risk;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureReviewSuppressService(
    ConnectSecurePatchService patchService,
    ConnectSecureSuppressService suppressService,
    ConnectSecureCacheService cache)
{
    private readonly Dictionary<int, IReadOnlyList<SuppressibleRemediationEntry>> _remediationCache = new();
    private readonly Dictionary<int, IReadOnlyList<SuppressibleProblemEntry>> _problemCache = new();

    public async Task<ReviewSuppressMatchResult> ResolveMatchAsync(
        int companyId,
        string product,
        string? cveIds,
        string? source,
        int? knownProblemId = null,
        CancellationToken ct = default)
    {
        var solutions = await GetRemediationsAsync(companyId, ct);
        var solutionMatcher = SuppressibleRemediationMatcher.Match(solutions, product);
        var solutionOptions = solutions.Select(ReviewSuppressEntry.FromSolution).ToList();

        if (UsesApplicationSolutionPath(product, source, cveIds))
            return BuildSolutionResult(solutionMatcher, solutionOptions);

        var cveIdList = GetCveIdsForSuppress(product, cveIds);
        var problemMatch = await ResolveProblemMatchAsync(
            companyId, product, source, knownProblemId, cveIdList, ct);

        IReadOnlyList<ReviewSuppressEntry> problemOptions = [];
        if (!problemMatch.HasMatch && !problemMatch.IsAmbiguous)
        {
            var problems = await GetProblemsAsync(companyId, ct);
            problemMatch = SuppressibleProblemMatcher.Match(problems, cveIdList);
            problemOptions = problems.Select(ReviewSuppressEntry.FromProblem).ToList();
        }
        else if (problemMatch.IsAmbiguous)
        {
            problemOptions = problemMatch.AmbiguousMatches.ToList();
        }

        var allOptions = problemOptions.Concat(solutionOptions).ToList();

        if (problemMatch.HasMatch)
            return problemMatch with { AllOptions = allOptions };

        if (problemMatch.IsAmbiguous)
            return problemMatch with { AllOptions = allOptions };

        if (solutionMatcher.HasMatch)
        {
            return new ReviewSuppressMatchResult(
                ReviewSuppressEntry.FromSolution(solutionMatcher.Entry!),
                [],
                allOptions);
        }

        if (solutionMatcher.IsAmbiguous)
        {
            return new ReviewSuppressMatchResult(
                null,
                solutionMatcher.AmbiguousMatches.Select(ReviewSuppressEntry.FromSolution).ToList(),
                allOptions);
        }

        return new ReviewSuppressMatchResult(null, [], allOptions);
    }

    public async Task<PatchOperationResult> SuppressCompanyWideAsync(
        int companyId,
        ReviewSuppressEntry entry,
        string product,
        string reason,
        string comments,
        CancellationToken ct = default)
    {
        var request = new SuppressVulnerabilityRequest
        {
            CompanyId = companyId,
            AssetId = 0,
            Product = entry.Kind == ReviewSuppressTargetKind.Problem
                ? entry.Label
                : (string.IsNullOrWhiteSpace(product) ? entry.Label : product),
            Reason = reason,
            Comments = comments
        };

        if (entry.Kind == ReviewSuppressTargetKind.Problem)
            request.ProblemId = entry.Id;
        else
            request.SolutionId = entry.Id;

        return await suppressService.SuppressAsync(request, ct);
    }

    public async Task<int?> ResolveSuppressRecordIdAsync(
        int companyId,
        int? storedRecordId,
        int? problemId,
        int? solutionId,
        CancellationToken ct = default)
    {
        if (storedRecordId is > 0)
            return storedRecordId;

        return await suppressService.FindSuppressRecordIdAsync(companyId, problemId, solutionId, ct);
    }

    public Task<PatchOperationResult> UnsuppressCompanyWideAsync(
        int suppressRecordId,
        CancellationToken ct = default) =>
        suppressService.UnsuppressAsync(suppressRecordId, ct);

    public void InvalidateCache(int companyId)
    {
        if (companyId <= 0)
            return;

        _remediationCache.Remove(companyId);
        _problemCache.Remove(companyId);
        cache.InvalidateCompany(companyId);
    }

    private static bool UsesApplicationSolutionPath(string product, string? source, string? cveIds) =>
        ReviewSuppressPathHelper.UsesApplicationSolutionPath(source, product, cveIds);

    private static ReviewSuppressMatchResult BuildSolutionResult(
        SuppressibleRemediationMatch matcher,
        IReadOnlyList<ReviewSuppressEntry> allOptions)
    {
        if (matcher.HasMatch)
            return new ReviewSuppressMatchResult(ReviewSuppressEntry.FromSolution(matcher.Entry!), [], allOptions);

        if (matcher.IsAmbiguous)
        {
            return new ReviewSuppressMatchResult(
                null,
                matcher.AmbiguousMatches.Select(ReviewSuppressEntry.FromSolution).ToList(),
                allOptions);
        }

        return new ReviewSuppressMatchResult(null, [], allOptions);
    }

    private async Task<ReviewSuppressMatchResult> ResolveProblemMatchAsync(
        int companyId,
        string product,
        string? source,
        int? knownProblemId,
        IReadOnlyList<string> cveIdList,
        CancellationToken ct)
    {
        if (knownProblemId is > 0)
        {
            var label = cveIdList.FirstOrDefault()
                ?? (CveReferenceHelper.IsCveOnlyProduct(product) ? product : product);
            return new ReviewSuppressMatchResult(
                new ReviewSuppressEntry(ReviewSuppressTargetKind.Problem, knownProblemId.Value, label, 0),
                [],
                []);
        }

        if (cveIdList.Count == 0)
            return new ReviewSuppressMatchResult(null, [], []);

        foreach (var cve in cveIdList)
        {
            var lookedUp = await patchService.LookupProblemByNameAsync(companyId, cve, source, ct: ct);
            if (lookedUp is not null)
                return new ReviewSuppressMatchResult(ReviewSuppressEntry.FromProblem(lookedUp), [], []);
        }

        return new ReviewSuppressMatchResult(null, [], []);
    }

    private async Task<IReadOnlyList<SuppressibleRemediationEntry>> GetRemediationsAsync(
        int companyId,
        CancellationToken ct)
    {
        if (_remediationCache.TryGetValue(companyId, out var cached))
            return cached;

        var entries = await patchService.GetSuppressibleRemediationsAsync(companyId, ct);
        _remediationCache[companyId] = entries;
        return entries;
    }

    private async Task<IReadOnlyList<SuppressibleProblemEntry>> GetProblemsAsync(
        int companyId,
        CancellationToken ct)
    {
        if (_problemCache.TryGetValue(companyId, out var cached))
            return cached;

        var entries = await patchService.GetSuppressibleProblemsAsync(companyId, ct);
        _problemCache[companyId] = entries;
        return entries;
    }

    private static IReadOnlyList<string> GetCveIdsForSuppress(string product, string? cveIds)
    {
        if (CveReferenceHelper.IsCveOnlyProduct(product))
            return CveReferenceHelper.SplitCveIds(product);

        return CveReferenceHelper.SplitCveIds(cveIds);
    }
}
