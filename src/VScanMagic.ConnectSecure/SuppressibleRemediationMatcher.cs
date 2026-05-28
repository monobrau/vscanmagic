using VScanMagic.Core.Risk;

namespace VScanMagic.ConnectSecure;

public sealed record SuppressibleRemediationMatch(
    SuppressibleRemediationEntry? Entry,
    IReadOnlyList<SuppressibleRemediationEntry> AmbiguousMatches)
{
    public bool HasMatch => Entry is not null;
    public bool IsAmbiguous => Entry is null && AmbiguousMatches.Count > 1;
}

public static class SuppressibleRemediationMatcher
{
    public static SuppressibleRemediationMatch Match(
        IReadOnlyList<SuppressibleRemediationEntry> entries,
        string product)
    {
        if (entries.Count == 0 || string.IsNullOrWhiteSpace(product))
            return new SuppressibleRemediationMatch(null, []);

        var candidates = new List<(SuppressibleRemediationEntry Entry, int Score)>();
        foreach (var name in ProductNameNormalizer.ParseProductNames(product))
        {
            var normalizedFinding = Normalize(name);
            if (normalizedFinding.Length == 0)
                continue;

            foreach (var entry in entries)
            {
                var score = ScoreMatch(normalizedFinding, Normalize(entry.Product));
                if (score > 0)
                    candidates.Add((entry, score));
            }
        }

        if (candidates.Count == 0)
            return new SuppressibleRemediationMatch(null, []);

        var bestScore = candidates.Max(c => c.Score);
        var best = candidates
            .Where(c => c.Score == bestScore)
            .Select(c => c.Entry)
            .DistinctBy(e => e.SolutionId)
            .ToList();

        if (best.Count == 1)
            return new SuppressibleRemediationMatch(best[0], []);

        return new SuppressibleRemediationMatch(null, best);
    }

    private static int ScoreMatch(string finding, string entry)
    {
        if (finding.Length == 0 || entry.Length == 0)
            return 0;

        if (finding.Equals(entry, StringComparison.OrdinalIgnoreCase))
            return 1000 + entry.Length;

        var findingMajor = ProductConsolidator.GetProductMajorVersion(finding);
        var entryMajor = ProductConsolidator.GetProductMajorVersion(entry);
        if (findingMajor.Equals(entryMajor, StringComparison.OrdinalIgnoreCase))
            return 850 + entryMajor.Length;

        var findingGroup = ProductConsolidator.GetTimeEstimateGroupKey(finding);
        var entryGroup = ProductConsolidator.GetTimeEstimateGroupKey(entry);
        if (findingGroup.Equals(entryGroup, StringComparison.OrdinalIgnoreCase))
            return 800 + entryGroup.Length;

        if (finding.StartsWith(entry, StringComparison.OrdinalIgnoreCase) ||
            entry.StartsWith(finding, StringComparison.OrdinalIgnoreCase))
            return 500 + Math.Min(finding.Length, entry.Length);

        if (finding.Contains(entry, StringComparison.OrdinalIgnoreCase) ||
            entry.Contains(finding, StringComparison.OrdinalIgnoreCase))
            return 100 + Math.Min(finding.Length, entry.Length);

        return 0;
    }

    private static string Normalize(string value) =>
        string.Join(' ', value.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries)).Trim();
}
