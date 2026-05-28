namespace VScanMagic.ConnectSecure;

using System.Text.Json;
using VScanMagic.Core.Services;

public static class PatchJobCorrelationHelper
{
    private static readonly TimeSpan MatchWindow = TimeSpan.FromMinutes(30);

    internal static bool IsPatchJobType(string? type)
    {
        if (string.IsNullOrWhiteSpace(type))
            return false;

        return type.Contains("patch", StringComparison.OrdinalIgnoreCase) ||
               type.Contains("remediation", StringComparison.OrdinalIgnoreCase) ||
               type.Contains("application", StringComparison.OrdinalIgnoreCase) ||
               type.Contains("linux", StringComparison.OrdinalIgnoreCase) ||
               type.Equals("os", StringComparison.OrdinalIgnoreCase);
    }

    internal static ParsedConnectSecureJob ParseConnectSecureJob(JsonElement row) =>
        ParsePatchJobViewRow(row);

    internal static ParsedConnectSecureJob ParsePatchJobViewRow(JsonElement row)
    {
        var updatedText = ConnectSecureJsonReader.GetString(row, "updated", "created");
        var updated = VScanMagic.Core.DisplayTime.ParseApiTimestamp(updatedText);

        var jobId = ConnectSecureJsonReader.GetString(row, "job_id", "jobId", "patch_id", "patchId") ?? "";
        var status = ConnectSecureJsonReader.GetString(row, "job_status", "jobStatus", "status") ?? "";
        var productName = ConnectSecureJsonReader.GetString(
            row,
            "product_name",
            "productName",
            "software_name",
            "softwareName",
            "application_name",
            "applicationName") ?? "";
        var type = ConnectSecureJsonReader.GetString(row, "type") ?? "";

        ParsePatchJobCounts(row, out var successCount, out var failedCount, out var pendingCount);
        var assetIds = ParsePatchJobAssetIds(row, out var hostNames);

        var description = BuildPatchJobDescription(productName, successCount, failedCount, pendingCount);
        var hostName = hostNames.Count switch
        {
            0 => null,
            1 => hostNames[0],
            _ => $"{hostNames[0]} +{hostNames.Count - 1} more"
        };

        var hostDetails = ParsePatchJobHostDetails(row);

        return new ParsedConnectSecureJob(
            jobId,
            type,
            status,
            description,
            hostName,
            AgentIp: null,
            AgentId: null,
            updated,
            ProductName: string.IsNullOrWhiteSpace(productName) ? null : productName,
            AssetIds: assetIds,
            SuccessCount: successCount,
            FailedCount: failedCount,
            PendingCount: pendingCount,
            HostDetails: hostDetails);
    }

    public static ParsedConnectSecureJob? FindBestMatch(
        PatchActivityEntry local,
        IReadOnlyList<ParsedConnectSecureJob> remoteJobs,
        IReadOnlySet<string> alreadyLinkedJobIds)
    {
        if (!string.IsNullOrWhiteSpace(local.ConnectSecureJobId))
        {
            var linked = remoteJobs.FirstOrDefault(job =>
                string.Equals(job.JobId, local.ConnectSecureJobId, StringComparison.OrdinalIgnoreCase));
            if (linked is not null && ProductNameMatches(local, linked))
                return linked;
        }

        ParsedConnectSecureJob? best = null;
        var bestScore = int.MinValue;

        foreach (var remote in remoteJobs)
        {
            if (string.IsNullOrWhiteSpace(remote.JobId) || alreadyLinkedJobIds.Contains(remote.JobId))
                continue;
            if (!IsPatchJobType(remote.Type))
                continue;

            var score = ScoreMatch(local, remote);
            if (score > bestScore)
            {
                bestScore = score;
                best = remote;
            }
        }

        if (best is null || bestScore < 3)
            return null;

        if (!ProductNameMatches(local, best))
            return null;

        return best;
    }

    internal static bool ProductNameMatches(PatchActivityEntry local, ParsedConnectSecureJob remote)
    {
        if (string.IsNullOrWhiteSpace(local.Product))
            return true;

        if (!string.IsNullOrWhiteSpace(remote.ProductName) &&
            string.Equals(local.Product, remote.ProductName, StringComparison.OrdinalIgnoreCase))
            return true;

        return remote.Description.Contains(local.Product, StringComparison.OrdinalIgnoreCase);
    }

    internal static int ScoreMatch(PatchActivityEntry local, ParsedConnectSecureJob remote)
    {
        var score = 0;
        var agentIds = local.AgentIds?.Where(id => id > 0).ToHashSet() ?? [];
        if (remote.AgentId is > 0 && agentIds.Contains(remote.AgentId.Value))
            score += 4;

        var assetIds = local.AssetIds?.Where(id => id > 0).ToHashSet() ?? [];
        if (remote.AssetIds is { Count: > 0 } && assetIds.Count > 0 &&
            remote.AssetIds.Any(assetIds.Contains))
            score += 5;

        if (remote.Updated is not null &&
            Math.Abs((remote.Updated.Value - local.RequestedAt).TotalMinutes) <= MatchWindow.TotalMinutes)
            score += 3;

        if (!string.IsNullOrWhiteSpace(local.HostName) &&
            !string.IsNullOrWhiteSpace(remote.HostName) &&
            HostNamesMatch(local.HostName, remote.HostName))
            score += 2;

        if (!string.IsNullOrWhiteSpace(local.Product) &&
            !string.IsNullOrWhiteSpace(remote.ProductName) &&
            string.Equals(local.Product, remote.ProductName, StringComparison.OrdinalIgnoreCase))
            score += 3;
        else if (!string.IsNullOrWhiteSpace(local.Product) &&
                 remote.Description.Contains(local.Product, StringComparison.OrdinalIgnoreCase))
            score += 2;

        if (local.IsOsPatch &&
            remote.Type.Contains("os", StringComparison.OrdinalIgnoreCase))
            score += 2;
        else if (!local.IsOsPatch &&
                 remote.Type.Contains("application", StringComparison.OrdinalIgnoreCase))
            score += 1;

        return score;
    }

    public static string ResolveVersionCheckStatus(PatchActivityEntry entry)
    {
        if (!string.IsNullOrWhiteSpace(entry.VersionCheckStatus))
            return entry.VersionCheckStatus;

        if (string.Equals(entry.Status, "Submitted", StringComparison.OrdinalIgnoreCase))
            return "";

        return entry.Status;
    }

    public static string FormatJobStatusLabel(string? status)
    {
        if (string.IsNullOrWhiteSpace(status))
            return "Unknown";

        return status.Trim();
    }

    public static bool IsTerminalJobStatus(string? status)
    {
        if (string.IsNullOrWhiteSpace(status))
            return false;

        var normalized = status.Trim();
        return normalized.Contains("success", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("completed", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("complete", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("fail", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("error", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("cancel", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("partial", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsInProgressJobStatus(string? status)
    {
        if (string.IsNullOrWhiteSpace(status))
            return false;

        var normalized = status.Trim();
        return normalized.Contains("progress", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("running", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("pending", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("initiated", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("queued", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("processing", StringComparison.OrdinalIgnoreCase);
    }

    internal static string BuildPatchJobDescription(
        string productName,
        int? successCount,
        int? failedCount,
        int? pendingCount)
    {
        if (successCount is null && failedCount is null && pendingCount is null)
            return productName;

        return $"{productName} — Success: {successCount ?? 0}, Failed: {failedCount ?? 0}, Pending: {pendingCount ?? 0}";
    }

    private static void ParsePatchJobCounts(
        JsonElement row,
        out int? successCount,
        out int? failedCount,
        out int? pendingCount)
    {
        successCount = null;
        failedCount = null;
        pendingCount = null;

        if (!row.TryGetProperty("msg", out var msg) || msg.ValueKind != JsonValueKind.Array)
            return;

        var values = msg.EnumerateArray().Select(ParseCountElement).ToList();
        if (values.Count >= 3)
        {
            successCount = values[0];
            failedCount = values[1];
            pendingCount = values[2];
        }
    }

    private static int ParseCountElement(JsonElement element) =>
        element.ValueKind switch
        {
            JsonValueKind.Number when element.TryGetInt32(out var n) => n,
            JsonValueKind.String when int.TryParse(element.GetString(), out var parsed) => parsed,
            _ => 0
        };

    private static IReadOnlyList<int> ParsePatchJobAssetIds(JsonElement row, out List<string> hostNames)
    {
        var details = ParsePatchJobHostDetails(row);
        hostNames = details
            .Select(d => d.HostName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToList();
        return details.Select(d => d.AssetId).Where(id => id > 0).ToList();
    }

    private static IReadOnlyList<PatchJobHostDetail> ParsePatchJobHostDetails(JsonElement row)
    {
        var results = new List<PatchJobHostDetail>();
        if (!row.TryGetProperty("patch_job_details", out var details) ||
            details.ValueKind != JsonValueKind.Object)
            return results;

        foreach (var property in details.EnumerateObject())
        {
            if (!int.TryParse(property.Name, out var assetId) || assetId <= 0)
                continue;

            var value = property.Value;
            results.Add(new PatchJobHostDetail(
                assetId,
                ConnectSecureJsonReader.GetString(value, "host_name", "hostName"),
                ConnectSecureJsonReader.GetString(value, "status"),
                ConnectSecureJsonReader.GetString(value, "status_msg", "statusMsg"),
                ConnectSecureJsonReader.GetString(value, "from_version", "fromVersion"),
                ConnectSecureJsonReader.GetString(value, "to_version", "toVersion")));
        }

        return results;
    }

    private static bool HostNamesMatch(string left, string right)
    {
        var a = NormalizeHost(left);
        var b = NormalizeHost(right);
        return string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ||
               a.StartsWith(b + ".", StringComparison.OrdinalIgnoreCase) ||
               b.StartsWith(a + ".", StringComparison.OrdinalIgnoreCase) ||
               left.Contains(b, StringComparison.OrdinalIgnoreCase) ||
               right.Contains(a, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeHost(string value)
    {
        var trimmed = value.Trim();
        var dot = trimmed.IndexOf('.');
        return dot > 0 ? trimmed[..dot] : trimmed;
    }

    public sealed record PatchJobHostDetail(
        int AssetId,
        string? HostName,
        string? Status,
        string? StatusMessage,
        string? FromVersion,
        string? ToVersion);

    public sealed record ParsedConnectSecureJob(
        string JobId,
        string Type,
        string Status,
        string Description,
        string? HostName,
        string? AgentIp,
        int? AgentId,
        DateTimeOffset? Updated,
        string? ProductName = null,
        IReadOnlyList<int>? AssetIds = null,
        int? SuccessCount = null,
        int? FailedCount = null,
        int? PendingCount = null,
        IReadOnlyList<PatchJobHostDetail>? HostDetails = null);
}
