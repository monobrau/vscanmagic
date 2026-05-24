using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureScanService(
    ConnectSecureClient client,
    ConnectSecureDiscoverySettingsService discoverySettings)
{
    public async Task<ScanTriggerResult> TriggerExternalScansAsync(
        int companyId,
        IReadOnlyList<int>? discoverySettingIds = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var ids = discoverySettingIds?.Where(id => id > 0).Distinct().ToList()
                  ?? await GetExternalDiscoverySettingIdsAsync(companyId, ct);

        if (ids.Count == 0)
            return new ScanTriggerResult(false, "No external scan targets configured for this company.", 0);

        await discoverySettings.TriggerExternalScanAsync(companyId, ids, ct);
        return new ScanTriggerResult(true, $"Queued external scan for {ids.Count} target(s).", ids.Count);
    }

    public async Task<ScanTriggerResult> TriggerInternalScansAsync(
        int companyId,
        IReadOnlyList<int>? discoverySettingIds = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var ids = discoverySettingIds?.Where(id => id > 0).Distinct().ToList()
                  ?? await GetInternalDiscoverySettingIdsAsync(companyId, ct);

        if (ids.Count == 0)
            return new ScanTriggerResult(false, "No internal subnet scan targets mapped for this company.", 0);

        var triggered = 0;
        foreach (var id in ids)
        {
            await TriggerInternalDiscoveryScanAsync(id, ct);
            triggered++;
        }

        return new ScanTriggerResult(
            true,
            $"Queued internal subnet scan for {triggered} target(s).",
            triggered);
    }

    public async Task<ScanTriggerResult> TriggerAgentUpdatesAsync(
        int companyId,
        IReadOnlyList<int>? agentIds = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            throw new ArgumentOutOfRangeException(nameof(companyId), "Company id is required.");

        var ids = agentIds?.Where(id => id > 0).Distinct().ToList()
                  ?? await GetLightweightAgentIdsAsync(companyId, ct);

        if (ids.Count == 0)
            return new ScanTriggerResult(false, "No agents found for this company.", 0);

        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/reset_agents",
            body: new Dictionary<string, object?>
            {
                ["company_id"] = companyId,
                ["agent_ids"] = ids.ToArray(),
                ["message"] = "update_agent"
            },
            ct: ct);

        var success = ConnectSecureJsonReader.GetBool(response, "status");
        var message = ConnectSecureJsonReader.GetString(response, "message");
        if (string.IsNullOrWhiteSpace(message))
            message = success
                ? $"Requested agent update for {ids.Count} agent(s)."
                : "ConnectSecure rejected the agent update request.";

        if (!success)
            throw new InvalidOperationException(message);

        return new ScanTriggerResult(true, message, ids.Count);
    }

    private async Task TriggerInternalDiscoveryScanAsync(int discoverySettingId, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            $"/r/company/discovery_settings/{discoverySettingId}",
            ct: ct);

        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Discovery setting {discoverySettingId} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        dataNode["scan_later"] = false;

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/company/discovery_settings",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = discoverySettingId },
            ct: ct);
    }

    private async Task<List<int>> GetExternalDiscoverySettingIdsAsync(int companyId, CancellationToken ct)
    {
        var rows = await FetchDiscoverySettingsAsync(companyId, ct);
        return rows
            .Where(ConnectSecureCompanyReviewService.IsExternalDiscoverySetting)
            .Select(row => ConnectSecureJsonReader.GetInt(row, "id") ?? 0)
            .Where(id => id > 0)
            .Distinct()
            .ToList();
    }

    private async Task<List<int>> GetInternalDiscoverySettingIdsAsync(int companyId, CancellationToken ct)
    {
        var rows = await FetchDiscoverySettingsAsync(companyId, ct);
        return rows
            .Where(row => !ConnectSecureCompanyReviewService.IsExternalDiscoverySetting(row))
            .Select(row => ConnectSecureJsonReader.GetInt(row, "id") ?? 0)
            .Where(id => id > 0)
            .Distinct()
            .ToList();
    }

    private async Task<List<int>> GetLightweightAgentIdsAsync(int companyId, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            "/r/report_queries/lightweight_assets",
            new Dictionary<string, string>
            {
                ["company_id"] = companyId.ToString(),
                ["limit"] = "5000",
                ["skip"] = "0"
            },
            ct: ct);

        return ConnectSecureJsonReader.ExtractDataArray(response)
            .Select(row => ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0)
            .Where(id => id > 0)
            .Distinct()
            .ToList();
    }

    private async Task<List<JsonElement>> FetchDiscoverySettingsAsync(int companyId, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            "/r/report_queries/discovery_settings",
            ConnectSecureCompanyReviewService.CompanyQuery(companyId, limit: 500, orderBy: "updated desc"),
            ct: ct);

        return ConnectSecureJsonReader.ExtractDataArray(response);
    }
}
