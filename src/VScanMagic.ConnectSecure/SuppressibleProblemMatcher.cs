using VScanMagic.Core.Risk;

namespace VScanMagic.ConnectSecure;

public static class SuppressibleProblemMatcher
{
    public static ReviewSuppressMatchResult Match(
        IReadOnlyList<SuppressibleProblemEntry> entries,
        IReadOnlyList<string> cveIds)
    {
        if (entries.Count == 0 || cveIds.Count == 0)
            return new ReviewSuppressMatchResult(null, [], []);

        var normalized = cveIds
            .SelectMany(CveReferenceHelper.SplitCveIds)
            .Select(c => c.ToUpperInvariant())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (normalized.Count == 0)
            return new ReviewSuppressMatchResult(null, [], []);

        var matches = entries
            .Where(entry => normalized.Contains(entry.ProblemName.Trim().ToUpperInvariant()))
            .Select(ReviewSuppressEntry.FromProblem)
            .DistinctBy(entry => entry.SelectionKey)
            .ToList();

        if (matches.Count == 0)
        {
            matches = entries
                .Where(entry => normalized.Any(c =>
                    entry.ProblemName.Contains(c, StringComparison.OrdinalIgnoreCase)))
                .Select(ReviewSuppressEntry.FromProblem)
                .DistinctBy(entry => entry.SelectionKey)
                .ToList();
        }

        if (matches.Count == 1)
            return new ReviewSuppressMatchResult(matches[0], [], []);

        if (matches.Count > 1)
            return new ReviewSuppressMatchResult(null, matches, []);

        return new ReviewSuppressMatchResult(null, [], []);
    }
}
