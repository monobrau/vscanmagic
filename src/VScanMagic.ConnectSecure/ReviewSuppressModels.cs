namespace VScanMagic.ConnectSecure;

public enum ReviewSuppressTargetKind
{
    Solution,
    Problem
}

public sealed record SuppressibleProblemEntry(
    int ProblemId,
    string ProblemName,
    int AffectedAssets);

public sealed record ReviewSuppressEntry(
    ReviewSuppressTargetKind Kind,
    int Id,
    string Label,
    int AffectedAssets)
{
    public string SelectionKey => $"{Kind}:{Id}";

    public static bool TryParseSelectionKey(string? key, out ReviewSuppressTargetKind kind, out int id)
    {
        kind = default;
        id = 0;
        if (string.IsNullOrWhiteSpace(key))
            return false;

        var parts = key.Split(':', 2);
        if (parts.Length != 2 || !int.TryParse(parts[1], out id))
            return false;

        return Enum.TryParse(parts[0], ignoreCase: true, out kind) && id > 0;
    }

    public static ReviewSuppressEntry FromSolution(SuppressibleRemediationEntry entry) =>
        new(ReviewSuppressTargetKind.Solution, entry.SolutionId, entry.Product, entry.AffectedAssets);

    public static ReviewSuppressEntry FromProblem(SuppressibleProblemEntry entry) =>
        new(ReviewSuppressTargetKind.Problem, entry.ProblemId, entry.ProblemName, entry.AffectedAssets);
}

public sealed record ReviewSuppressMatchResult(
    ReviewSuppressEntry? Entry,
    IReadOnlyList<ReviewSuppressEntry> AmbiguousMatches,
    IReadOnlyList<ReviewSuppressEntry> AllOptions)
{
    public bool HasMatch => Entry is not null;
    public bool IsAmbiguous => Entry is null && AmbiguousMatches.Count > 1;
}
