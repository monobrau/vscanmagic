using System.Text.Json;
using VScanMagic.Core.Risk;

namespace VScanMagic.ConnectSecure;

internal static class ConnectSecureSuppressLookup
{
    public static bool ProblemNameMatches(string? candidate, string target)
    {
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(target))
            return false;

        if (candidate.Equals(target, StringComparison.OrdinalIgnoreCase))
            return true;

        return CveReferenceHelper.SplitCveIds(candidate)
            .Any(c => c.Equals(target, StringComparison.OrdinalIgnoreCase));
    }

    public static SuppressibleProblemEntry? FindProblemMatch(
        IEnumerable<JsonElement> rows,
        int companyId,
        string problemName,
        Func<JsonElement, int, bool> matchesCompany,
        Func<JsonElement, SuppressibleProblemEntry> parse)
    {
        foreach (var row in rows)
        {
            if (!matchesCompany(row, companyId))
                continue;

            var name = ConnectSecureJsonReader.GetString(row, "problem_name", "problemName", "name");
            if (string.IsNullOrWhiteSpace(name))
                name = ConnectSecureJsonReader.GetString(row, "software_name", "softwareName");

            if (!ProblemNameMatches(name, problemName))
                continue;

            var entry = parse(row);
            if (entry.ProblemId > 0 && !string.IsNullOrWhiteSpace(entry.ProblemName))
                return entry;
        }

        return null;
    }

    public static void MergeProblem(
        IDictionary<string, SuppressibleProblemEntry> merged,
        SuppressibleProblemEntry entry)
    {
        if (entry.ProblemId <= 0 || string.IsNullOrWhiteSpace(entry.ProblemName))
            return;

        var key = entry.ProblemName.Trim().ToUpperInvariant();
        if (!merged.TryGetValue(key, out var existing) || entry.AffectedAssets > existing.AffectedAssets)
            merged[key] = entry;
    }
}
