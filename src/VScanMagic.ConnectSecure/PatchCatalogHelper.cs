using VScanMagic.Core.Services;

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
        SplitVersionTokens(string.Join(' ', versions.Where(v => !string.IsNullOrWhiteSpace(v))));

    public static IReadOnlyList<string> SplitVersionTokens(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return [];

        return text
            .Split([' ', ',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static bool ProductNamesMatch(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return false;

        return string.Equals(left.Trim(), right.Trim(), StringComparison.OrdinalIgnoreCase);
    }

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
        IsEndOfLifeProductFix(fix);

    public static bool IsEndOfLifeProductFix(string? fix) =>
        !string.IsNullOrWhiteSpace(fix) &&
        fix.Contains("end of life", StringComparison.OrdinalIgnoreCase);

    public static bool MeetsSeverityFilter(string? severity, string filter) =>
        SeverityRank(severity) >= MinimumSeverityRank(filter);

    public static int MinimumSeverityRank(string filter) =>
        filter switch
        {
            "critical" => SeverityRank("Critical"),
            "high+" => SeverityRank("High"),
            _ => 0
        };

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
        if (installedVersions.Count > 0)
            return HostPatchStatus.Pending;
        if (!IsVersionComparableTarget(targetFix))
            return HostPatchStatus.Pending;
        return HostPatchStatus.Unknown;
    }

    /// <summary>
    /// CS sometimes returns qualitative fix text (e.g. "Latest Patch") instead of a semver.
    /// Those targets cannot be version-checked but online remediation hosts are still patchable.
    /// </summary>
    public static bool IsVersionComparableTarget(string? targetFix)
    {
        if (string.IsNullOrWhiteSpace(targetFix) || IsEndOfLifeFix(targetFix))
            return false;

        var trimmed = targetFix.Trim();
        if (trimmed.Equals("Latest Patch", StringComparison.OrdinalIgnoreCase))
            return false;

        return System.Text.RegularExpressions.Regex.IsMatch(
            trimmed,
            @"(\d+(?:\.\d+)+|KB?\d{4,})",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    public static bool IsOsUpdateTarget(string? targetFix) =>
        !string.IsNullOrWhiteSpace(targetFix) &&
        System.Text.RegularExpressions.Regex.IsMatch(
            targetFix.Trim(),
            @"^(KB)?\d+$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    public static string NormalizeOsUpdateFix(string? fix)
    {
        if (string.IsNullOrWhiteSpace(fix))
            return "";

        var trimmed = fix.Trim();
        if (trimmed.StartsWith("KB", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[2..].Trim();

        return trimmed;
    }

    public static IReadOnlySet<int> BuildOsPendingAssetIndex(
        IEnumerable<OsPendingPatchEntry> pendingPatches,
        string targetFix)
    {
        var normalizedFix = NormalizeOsUpdateFix(targetFix);
        if (string.IsNullOrWhiteSpace(normalizedFix))
            return new HashSet<int>();

        var assetIds = new HashSet<int>();
        foreach (var entry in pendingPatches)
        {
            if (!string.Equals(NormalizeOsUpdateFix(entry.Fix), normalizedFix, StringComparison.OrdinalIgnoreCase))
                continue;

            foreach (var assetId in entry.AssetIds.Where(id => id > 0))
                assetIds.Add(assetId);
        }

        return assetIds;
    }

    public static HostPatchStatus DetermineOsHostStatus(
        bool online,
        PatchAssetDetail detail,
        string targetFix,
        IReadOnlySet<int> pendingAssetIds)
    {
        if (!online)
            return HostPatchStatus.Offline;

        if (!IsOsUpdateTarget(targetFix))
            return DetermineHostStatus(online, isEndOfLife: false, detail.Versions, targetFix);

        var assetIds = new[] { detail.AssetId, detail.AgentId }.Where(id => id > 0);
        var stillPending = assetIds.Any(pendingAssetIds.Contains);
        return stillPending ? HostPatchStatus.Pending : HostPatchStatus.AtTarget;
    }

    public static IReadOnlyList<PatchHostView> BuildOsHostViews(
        IEnumerable<PatchAssetDetail> details,
        string targetFix,
        IReadOnlySet<int> pendingAssetIds) =>
        details
            .Select(detail =>
            {
                var status = DetermineOsHostStatus(detail.OnlineStatus, detail, targetFix, pendingAssetIds);
                return new PatchHostView(detail, status, StatusLabel(status));
            })
            .ToList();

    public static string StatusLabel(HostPatchStatus status) =>
        status switch
        {
            HostPatchStatus.Offline => "Offline",
            HostPatchStatus.EndOfLife => "EOL",
            HostPatchStatus.AtTarget => "At target",
            HostPatchStatus.Pending => "Pending patch",
            HostPatchStatus.Unknown => "Version unknown",
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

    public static PatchVerificationResult BuildVerificationResult(
        string jobId,
        IReadOnlyList<int> targetedAgentIds,
        IEnumerable<PatchHostView> hostViews,
        DateTimeOffset? verifiedAt = null)
    {
        var agentSet = targetedAgentIds.Where(id => id > 0).ToHashSet();
        var matched = hostViews
            .Where(view => agentSet.Contains(view.Detail.AgentId))
            .ToList();

        var hostResults = matched
            .Select(view => new PatchVerificationHostResult(
                view.Detail.AgentId,
                view.Detail.HostName,
                view.Status,
                FormatVersionSummary(view.Detail.Versions),
                view.StatusLabel))
            .OrderBy(result => result.HostName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var atTarget = hostResults.Count(r => r.Status == HostPatchStatus.AtTarget);
        var pending = hostResults.Count(r => r.Status == HostPatchStatus.Pending);
        var offline = hostResults.Count(r => r.Status == HostPatchStatus.Offline);
        var endOfLife = hostResults.Count(r => r.Status == HostPatchStatus.EndOfLife);
        var unknown = hostResults.Count(r => r.Status == HostPatchStatus.Unknown);
        var total = hostResults.Count;
        var verified = verifiedAt ?? DateTimeOffset.Now;

        var status = DetermineVerificationStatus(atTarget, pending, offline, endOfLife, unknown, total);
        var summary = FormatVerificationSummary(status, atTarget, pending, offline, endOfLife, unknown, total);

        return new PatchVerificationResult(
            jobId,
            status,
            atTarget,
            pending,
            offline,
            endOfLife,
            unknown,
            total,
            summary,
            hostResults,
            verified,
            ConnectSecureInsight: null);
    }

    public static string BuildConnectSecureJobInsight(
        PatchActivityEntry entry,
        PatchJobCorrelationHelper.ParsedConnectSecureJob? remote)
    {
        if (remote is null)
        {
            return entry.RequestedAt < DateTimeOffset.Now.AddHours(-1)
                ? "No matching ConnectSecure patch job found in patch_jobview. Check the CS portal Patch Jobs tab for this product and time."
                : "ConnectSecure patch job not visible yet in patch_jobview — use Refresh jobs in a minute.";
        }

        var parts = new List<string>
        {
            $"CS job {remote.JobId}: {PatchJobCorrelationHelper.FormatJobStatusLabel(remote.Status)}"
        };

        if (remote.SuccessCount is not null || remote.FailedCount is not null || remote.PendingCount is not null)
            parts.Add($"agent counts Success {remote.SuccessCount ?? 0}, Failed {remote.FailedCount ?? 0}, Pending {remote.PendingCount ?? 0}");

        if (remote.HostDetails is { Count: > 0 })
        {
            foreach (var host in remote.HostDetails)
            {
                var hostLabel = string.IsNullOrWhiteSpace(host.HostName) ? $"asset {host.AssetId}" : host.HostName;
                var version = string.IsNullOrWhiteSpace(host.FromVersion) && string.IsNullOrWhiteSpace(host.ToVersion)
                    ? ""
                    : $" ({host.FromVersion ?? "?"} → {host.ToVersion ?? "?"})";
                parts.Add($"{hostLabel}: {host.Status ?? "Unknown"}{version}");
            }
        }

        if (PatchJobCorrelationHelper.IsInProgressJobStatus(remote.Status) &&
            entry.RequestedAt < DateTimeOffset.Now.AddHours(-2))
        {
            parts.Add(
                "Job has been in progress for 2+ hours. A vulnerability scan does not execute patches — the ConnectSecure agent must run the patch job. Check agent type (probe vs lightweight), Patch Management settings, and the job in the CS portal.");
        }
        else if (PatchJobCorrelationHelper.IsTerminalJobStatus(remote.Status) &&
                 (remote.SuccessCount ?? 0) > 0)
        {
            parts.Add(
                "ConnectSecure reports the patch job finished successfully. Version check uses remediation inventory — Verify queues an inventory scan; re-run Verify after agents finish if versions are still pending.");
        }

        return string.Join(". ", parts);
    }

    public static bool ShouldInferRemediationCleared(
        PatchJobCorrelationHelper.ParsedConnectSecureJob? remoteJob,
        int targetedHostCount,
        int matchedHostCount)
    {
        if (remoteJob is null || targetedHostCount <= 0 || matchedHostCount > 0)
            return false;

        if ((remoteJob.SuccessCount ?? 0) <= 0)
            return false;

        var status = remoteJob.Status ?? "";
        return status.Contains("success", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("partial", StringComparison.OrdinalIgnoreCase);
    }

    public static PatchVerificationResult BuildRemediationClearedVerificationResult(
        string jobId,
        IReadOnlyList<int> targetedAgentIds,
        DateTimeOffset? verifiedAt = null)
    {
        var verified = verifiedAt ?? DateTimeOffset.Now;
        return new PatchVerificationResult(
            jobId,
            "Verified",
            targetedAgentIds.Count,
            0,
            0,
            0,
            0,
            targetedAgentIds.Count,
            $"Verified: product no longer appears in the remediation plan for {targetedAgentIds.Count} patched host(s) (typical after a successful patch).",
            [],
            verified,
            ConnectSecureInsight: null);
    }

    public static string DetermineVerificationStatus(
        int atTarget,
        int pending,
        int offline,
        int endOfLife,
        int unknown,
        int total)
    {
        if (total == 0)
            return "Unverifiable";

        var verifiableOnline = total - offline;
        if (verifiableOnline <= 0)
            return "Pending verification";

        if (atTarget == verifiableOnline)
            return "Verified";

        if (atTarget > 0)
            return "Partial";

        if (pending > 0 || unknown > 0)
            return "Pending verification";

        if (endOfLife > 0 && atTarget == 0 && pending == 0)
            return "End of life";

        return "Failed";
    }

    public static string FormatVerificationSummary(
        string status,
        int atTarget,
        int pending,
        int offline,
        int endOfLife,
        int unknown,
        int total)
    {
        if (total == 0)
            return "No matching hosts found to verify.";

        var parts = new List<string> { $"{atTarget}/{total} at target" };
        if (pending > 0)
            parts.Add($"{pending} pending");
        if (offline > 0)
            parts.Add($"{offline} offline");
        if (endOfLife > 0)
            parts.Add($"{endOfLife} EOL");
        if (unknown > 0)
            parts.Add($"{unknown} unknown");

        return $"{status}: {string.Join(", ", parts)}.";
    }

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

            if (VersionTokensMatch(normalized, target))
                return true;

            if (targetFix.Contains(normalized, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static bool VersionTokensMatch(string installed, string target)
    {
        if (installed.Equals(target, StringComparison.OrdinalIgnoreCase))
            return true;

        if (installed.StartsWith(target + ".", StringComparison.OrdinalIgnoreCase) ||
            target.StartsWith(installed + ".", StringComparison.OrdinalIgnoreCase))
            return true;

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
