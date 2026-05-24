namespace VScanMagic.ConnectSecure;

public sealed record PatchableProductGroup(
    string Product,
    string Severity,
    string TargetFix,
    bool IsEndOfLife,
    int ReportedAssets,
    IReadOnlyList<int> SolutionIds,
    int RemediationEntryCount);

public static class PatchCatalogHelper
{
    private static readonly Dictionary<string, int> SeverityOrder = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Critical"] = 4,
        ["High"] = 3,
        ["Medium"] = 2,
        ["Low"] = 1
    };

    public static IReadOnlyList<PatchableProductGroup> GroupByProduct(IEnumerable<PatchableApplicationEntry> entries)
    {
        return entries
            .GroupBy(entry => entry.Product.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var rows = group
                    .OrderByDescending(entry => SeverityRank(entry.Severity))
                    .ThenByDescending(entry => entry.AffectedAssets)
                    .ToList();
                var primary = rows[0];
                var fix = PickDisplayFix(rows);
                return new PatchableProductGroup(
                    primary.Product,
                    rows.MaxBy(entry => SeverityRank(entry.Severity))?.Severity ?? primary.Severity,
                    fix,
                    IsEndOfLifeFix(fix),
                    rows.Max(entry => entry.AffectedAssets),
                    rows.Select(entry => entry.SolutionId).Distinct().ToList(),
                    rows.Count);
            })
            .OrderByDescending(group => SeverityRank(group.Severity))
            .ThenByDescending(group => group.ReportedAssets)
            .ThenBy(group => group.Product, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static IReadOnlyList<PatchAssetDetail> MergeAssetDetails(IEnumerable<PatchAssetDetail> details)
    {
        var merged = new Dictionary<string, PatchAssetDetail>(StringComparer.Ordinal);
        foreach (var detail in details)
        {
            var key = MergeAssetKey(detail);
            if (string.IsNullOrEmpty(key))
                continue;

            if (!merged.TryGetValue(key, out var existing))
            {
                merged[key] = detail with { Versions = DistinctVersions(detail.Versions) };
                continue;
            }

            merged[key] = existing with
            {
                AssetId = existing.AssetId > 0 ? existing.AssetId : detail.AssetId,
                AgentId = existing.AgentId > 0 ? existing.AgentId : detail.AgentId,
                OnlineStatus = existing.OnlineStatus || detail.OnlineStatus,
                Versions = DistinctVersions(existing.Versions.Concat(detail.Versions))
            };
        }

        return merged.Values
            .OrderByDescending(detail => detail.OnlineStatus)
            .ThenBy(detail => detail.HostName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(detail => detail.Ip, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static string MergeAssetKey(PatchAssetDetail detail)
    {
        if (detail.AgentId > 0 && detail.AssetId > 0)
            return $"a{detail.AssetId}:g{detail.AgentId}";
        if (detail.AssetId > 0)
            return $"a{detail.AssetId}";
        if (detail.AgentId > 0)
            return $"g{detail.AgentId}";
        return "";
    }

    public static string FormatVersionSummary(IReadOnlyList<string> versions)
    {
        var distinct = DistinctVersions(versions);
        return distinct.Count == 0 ? "—" : string.Join(", ", distinct.Take(3));
    }

    public static IReadOnlyList<string> NormalizeVersions(IEnumerable<string> versions) =>
        versions
            .Where(version => !string.IsNullOrWhiteSpace(version))
            .Select(version => version.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    public static string FormatFixSummary(string fix)
    {
        if (string.IsNullOrWhiteSpace(fix))
            return "—";
        if (IsEndOfLifeFix(fix))
            return "End of life (see CS guidance)";

        var trimmed = fix.Trim();
        if (trimmed.Length <= 48)
            return trimmed;

        return trimmed[..48] + "…";
    }

    public static int SeverityRank(string? severity) =>
        severity is not null && SeverityOrder.TryGetValue(severity.Trim(), out var rank) ? rank : 0;

    private static string PickDisplayFix(IReadOnlyList<PatchableApplicationEntry> rows)
    {
        foreach (var row in rows)
        {
            if (!IsEndOfLifeFix(row.Fix) && !string.IsNullOrWhiteSpace(row.Fix))
                return row.Fix.Trim();
        }

        return rows[0].Fix.Trim();
    }

    private static bool IsEndOfLifeFix(string? fix) =>
        !string.IsNullOrWhiteSpace(fix) &&
        fix.Contains("end of life", StringComparison.OrdinalIgnoreCase);

    private static IReadOnlyList<string> DistinctVersions(IEnumerable<string> versions) =>
        NormalizeVersions(versions);

    public static HostPatchStatus DetermineHostStatus(
        bool online,
        bool isEndOfLife,
        IReadOnlyList<string> installedVersions,
        string targetFix)
    {
        if (!online)
            return HostPatchStatus.Offline;
        if (isEndOfLife || IsEndOfLifeFix(targetFix))
            return HostPatchStatus.EndOfLife;
        if (IsAtTargetVersion(installedVersions, targetFix))
            return HostPatchStatus.AtTarget;
        if (installedVersions.Count > 0 && !string.IsNullOrWhiteSpace(targetFix))
            return HostPatchStatus.Pending;
        return HostPatchStatus.Unknown;
    }

    public static string StatusLabel(HostPatchStatus status) =>
        status switch
        {
            HostPatchStatus.Offline => "Offline",
            HostPatchStatus.EndOfLife => "EOL",
            HostPatchStatus.AtTarget => "At target",
            HostPatchStatus.Pending => "Pending patch",
            _ => "Unknown"
        };

    public static IReadOnlyList<PatchHostView> BuildHostViews(
        IEnumerable<PatchAssetDetail> details,
        string targetFix,
        bool isEndOfLife) =>
        details
            .Select(detail =>
            {
                var status = DetermineHostStatus(
                    detail.OnlineStatus,
                    isEndOfLife,
                    detail.Versions,
                    targetFix);
                return new PatchHostView(detail, status, StatusLabel(status));
            })
            .ToList();

    internal static bool IsAtTargetVersion(IReadOnlyList<string> installedVersions, string targetFix)
    {
        if (installedVersions.Count == 0 || string.IsNullOrWhiteSpace(targetFix))
            return false;

        var target = NormalizeVersionToken(targetFix);
        if (string.IsNullOrWhiteSpace(target))
            return false;

        foreach (var installed in installedVersions)
        {
            var normalized = NormalizeVersionToken(installed);
            if (string.IsNullOrWhiteSpace(normalized))
                continue;

            if (normalized.Equals(target, StringComparison.OrdinalIgnoreCase))
                return true;

            if (targetFix.Contains(normalized, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static string NormalizeVersionToken(string value)
    {
        var trimmed = value.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
            return "";

        var match = System.Text.RegularExpressions.Regex.Match(
            trimmed,
            @"\d+(?:\.\d+){1,4}");
        return match.Success ? match.Value : trimmed;
    }
}
