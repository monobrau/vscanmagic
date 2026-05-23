using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureDiscoverySettingsService(ConnectSecureClient client)
{
    private const string ExternalScanType = "externalscan";
    private const string DefaultInternalScanType = "networkscan";

    public async Task<int> CreateExternalScanTargetAsync(
        int companyId,
        string name,
        string addressInput,
        CancellationToken ct = default)
    {
        var validated = ExternalSubnetHelper.ParseAndValidateScanInput(addressInput);
        if (!validated.IsValid)
            throw new InvalidOperationException(string.Join(" ", validated.Errors));

        var data = BuildScanTargetData(companyId, name, validated, ExternalScanType, null);
        var id = await CreateDiscoverySettingAsync(data, ct);
        await TriggerExternalScanAsync(companyId, [id], ct);
        return id;
    }

    public async Task UpdateExternalScanTargetAsync(
        int discoverySettingId,
        int companyId,
        string name,
        string addressInput,
        CancellationToken ct = default)
    {
        var validated = ExternalSubnetHelper.ParseAndValidateScanInput(addressInput);
        if (!validated.IsValid)
            throw new InvalidOperationException(string.Join(" ", validated.Errors));

        await UpdateDiscoverySettingAsync(
            discoverySettingId,
            companyId,
            name,
            validated,
            ExternalScanType,
            null,
            ct);

        await TriggerExternalScanAsync(companyId, [discoverySettingId], ct);
    }

    public Task DeleteExternalScanTargetAsync(int discoverySettingId, CancellationToken ct = default) =>
        DeleteDiscoverySettingAsync(discoverySettingId, ct);

    public async Task<(int DiscoverySettingId, int MappingId)> CreateInternalSubnetAsync(
        int companyId,
        int probeAgentId,
        string name,
        string addressInput,
        CancellationToken ct = default)
    {
        if (probeAgentId <= 0)
            throw new InvalidOperationException("Select a probe agent for internal subnet scanning.");

        var validated = ExternalSubnetHelper.ParseAndValidateScanInput(addressInput);
        if (!validated.IsValid)
            throw new InvalidOperationException(string.Join(" ", validated.Errors));

        var template = await ResolveInternalDiscoveryTemplateAsync(companyId, ct);
        var discoverySettingId = await CreateDiscoverySettingAsync(
            BuildScanTargetData(companyId, name, validated, template.DiscoverySettingsType, template.AddressType),
            ct);
        var mappingId = await CreateAgentMappingAsync(companyId, probeAgentId, discoverySettingId, ct);
        return (discoverySettingId, mappingId);
    }

    public async Task UpdateInternalSubnetAsync(
        int discoverySettingId,
        int? mappingId,
        int companyId,
        int probeAgentId,
        string name,
        string addressInput,
        CancellationToken ct = default)
    {
        if (probeAgentId <= 0)
            throw new InvalidOperationException("Select a probe agent for internal subnet scanning.");

        var validated = ExternalSubnetHelper.ParseAndValidateScanInput(addressInput);
        if (!validated.IsValid)
            throw new InvalidOperationException(string.Join(" ", validated.Errors));

        var template = await ResolveInternalDiscoveryTemplateAsync(companyId, ct);
        await UpdateDiscoverySettingAsync(
            discoverySettingId,
            companyId,
            name,
            validated,
            template.DiscoverySettingsType,
            template.AddressType,
            ct);

        if (mappingId is > 0)
            await UpdateAgentMappingAsync(mappingId.Value, companyId, probeAgentId, discoverySettingId, ct);
        else
            await CreateAgentMappingAsync(companyId, probeAgentId, discoverySettingId, ct);
    }

    public async Task DeleteInternalSubnetAsync(int discoverySettingId, int? mappingId, CancellationToken ct = default)
    {
        if (mappingId is > 0)
        {
            try
            {
                await client.InvokeAuthenticatedAsync(
                    HttpMethod.Delete,
                    $"/d/company/agent_discoverysettings_mapping/{mappingId.Value}",
                    ct: ct);
            }
            catch (HttpRequestException)
            {
                // Mapping may already be gone; still attempt discovery setting delete.
            }
        }

        await DeleteDiscoverySettingAsync(discoverySettingId, ct);
    }

    public Task TriggerExternalScanAsync(int companyId, IReadOnlyList<int> discoverySettingIds, CancellationToken ct = default)
    {
        if (discoverySettingIds.Count == 0)
            return Task.CompletedTask;

        return client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/external_scan",
            body: new Dictionary<string, object?>
            {
                ["company_id"] = companyId,
                ["discovery_settings"] = discoverySettingIds.ToArray()
            },
            ct: ct);
    }

    private async Task<int> CreateDiscoverySettingAsync(Dictionary<string, object?> data, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/discovery_settings",
            body: new Dictionary<string, object?> { ["data"] = data },
            ct: ct);

        var id = ExtractCreatedId(response);
        if (id is null or <= 0)
            throw new InvalidOperationException("ConnectSecure did not return a discovery setting id.");

        return id.Value;
    }

    private async Task UpdateDiscoverySettingAsync(
        int discoverySettingId,
        int companyId,
        string name,
        ExternalSubnetHelper.ExternalScanTargetValidationResult validated,
        string discoverySettingsType,
        string? addressType,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            $"/r/company/discovery_settings/{discoverySettingId}",
            ct: ct);

        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Discovery setting {discoverySettingId} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        ApplyScanTargetFields(dataNode, companyId, name, validated, discoverySettingsType, addressType);

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/company/discovery_settings",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = discoverySettingId },
            ct: ct);
    }

    private Task DeleteDiscoverySettingAsync(int discoverySettingId, CancellationToken ct) =>
        client.InvokeAuthenticatedAsync(HttpMethod.Delete, $"/d/company/discovery_settings/{discoverySettingId}", ct: ct);

    private async Task<int> CreateAgentMappingAsync(
        int companyId,
        int probeAgentId,
        int discoverySettingId,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/agent_discoverysettings_mapping",
            body: new Dictionary<string, object?>
            {
                ["data"] = new Dictionary<string, object?>
                {
                    ["company_id"] = companyId,
                    ["agent_id"] = probeAgentId,
                    ["discovery_settings_id"] = discoverySettingId
                }
            },
            ct: ct);

        var id = ExtractCreatedId(response);
        if (id is null or <= 0)
            throw new InvalidOperationException("ConnectSecure did not return an agent mapping id.");

        return id.Value;
    }

    private async Task UpdateAgentMappingAsync(
        int mappingId,
        int companyId,
        int probeAgentId,
        int discoverySettingId,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            $"/r/company/agent_discoverysettings_mapping/{mappingId}",
            ct: ct);

        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Agent mapping {mappingId} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        dataNode["company_id"] = companyId;
        dataNode["agent_id"] = probeAgentId;
        dataNode["discovery_settings_id"] = discoverySettingId;

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/company/agent_discoverysettings_mapping",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = mappingId },
            ct: ct);
    }

    private async Task<InternalDiscoveryTemplate> ResolveInternalDiscoveryTemplateAsync(int companyId, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            "/r/report_queries/discovery_settings",
            ConnectSecureCompanyReviewService.CompanyQuery(companyId, limit: 500, orderBy: "updated desc"),
            ct: ct);

        foreach (var ds in ConnectSecureJsonReader.ExtractDataArray(response))
        {
            if (ConnectSecureCompanyReviewService.IsExternalDiscoverySetting(ds))
                continue;

            var type = ConnectSecureJsonReader.GetString(ds, "discovery_settings_type", "type");
            if (string.IsNullOrWhiteSpace(type))
                continue;

            return new InternalDiscoveryTemplate(
                type,
                ConnectSecureJsonReader.GetString(ds, "address_type", "addressType"));
        }

        return new InternalDiscoveryTemplate(DefaultInternalScanType, null);
    }

    private static Dictionary<string, object?> BuildScanTargetData(
        int companyId,
        string name,
        ExternalSubnetHelper.ExternalScanTargetValidationResult validated,
        string discoverySettingsType,
        string? addressType) =>
        ApplyScanTargetFields(new Dictionary<string, object?>(), companyId, name, validated, discoverySettingsType, addressType);

    private static Dictionary<string, object?> ApplyScanTargetFields(
        IDictionary<string, object?> target,
        int companyId,
        string name,
        ExternalSubnetHelper.ExternalScanTargetValidationResult validated,
        string discoverySettingsType,
        string? addressType)
    {
        target["company_id"] = companyId;
        target["name"] = name.Trim();
        target["discovery_settings_type"] = discoverySettingsType;
        target["address"] = validated.Address;
        target["target_ip"] = validated.TargetIp;
        target["scan_later"] = false;
        target["is_excluded"] = false;
        if (!string.IsNullOrWhiteSpace(addressType))
            target["address_type"] = addressType;

        return target is Dictionary<string, object?> dict ? dict : new Dictionary<string, object?>(target);
    }

    private static JsonObject ApplyScanTargetFields(
        JsonObject target,
        int companyId,
        string name,
        ExternalSubnetHelper.ExternalScanTargetValidationResult validated,
        string discoverySettingsType,
        string? addressType)
    {
        target["company_id"] = companyId;
        target["name"] = name.Trim();
        target["discovery_settings_type"] = discoverySettingsType;
        target["address"] = validated.Address;
        target["target_ip"] = validated.TargetIp;
        target["scan_later"] = false;
        target["is_excluded"] = false;
        if (!string.IsNullOrWhiteSpace(addressType))
            target["address_type"] = addressType;
        return target;
    }

    private static int? ExtractCreatedId(JsonElement response)
    {
        var id = ConnectSecureJsonReader.GetInt(response, "id");
        if (id is > 0)
            return id;

        if (response.TryGetProperty("data", out var dataEl))
        {
            id = ConnectSecureJsonReader.GetInt(dataEl, "id");
            if (id is > 0)
                return id;
        }

        if (response.TryGetProperty("id", out var idEl) && idEl.ValueKind == JsonValueKind.String &&
            int.TryParse(idEl.GetString(), out var parsed))
            return parsed;

        return null;
    }

    private sealed record InternalDiscoveryTemplate(string DiscoverySettingsType, string? AddressType);
}
