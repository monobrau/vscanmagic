using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureCompanyCredentialsService(ConnectSecureClient client)
{
    public async Task<IReadOnlyList<CompanyCredentialEntry>> ListCredentialsAsync(int companyId, CancellationToken ct = default)
    {
        var credentials = await FetchArrayAsync("/r/company/credentials", CompanyQuery(companyId), ct);
        var mappings = await ListCredentialMappingsAsync(companyId, ct);
        var agentsByCredential = mappings
            .GroupBy(m => m.CredentialsId)
            .ToDictionary(g => g.Key, g => (IReadOnlyList<int>)g.Select(x => x.AgentId).Distinct().ToList());

        return credentials
            .Select(c =>
            {
                var id = ConnectSecureJsonReader.GetInt(c, "id") ?? 0;
                c.TryGetProperty("params", out var paramsEl);
                agentsByCredential.TryGetValue(id, out var mappedAgents);
                return new CompanyCredentialEntry(
                    id,
                    ConnectSecureJsonReader.GetString(c, "name"),
                    ConnectSecureJsonReader.GetString(c, "credential_type", "credentialType"),
                    ConnectSecureJsonReader.GetString(c, "os_type", "osType"),
                    ConnectSecureJsonReader.GetString(c, "address_type", "addressType"),
                    ConnectSecureJsonReader.GetString(c, "address"),
                    ConnectSecureJsonReader.GetBool(c, "is_valid", "isValid"),
                    ConnectSecureJsonReader.GetString(c, "failure_reason", "failureReason"),
                    ConnectSecureParamsHelper.SummarizeParams(paramsEl),
                    mappedAgents ?? []);
            })
            .Where(c => c.Id > 0)
            .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<(CompanyCredentialEntry Entry, string ParamsJson)> GetCredentialAsync(int credentialId, int companyId, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            $"/r/company/credentials/{credentialId}",
            ct: ct);

        if (!response.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Credential {credentialId} was not found.");

        var mappings = await ListCredentialMappingsAsync(companyId, ct);
        var mappedAgents = mappings.Where(m => m.CredentialsId == credentialId).Select(m => m.AgentId).Distinct().ToList();
        data.TryGetProperty("params", out var paramsEl);
        var paramsJson = paramsEl.ValueKind == JsonValueKind.Object ? paramsEl.GetRawText() : "{}";

        var entry = new CompanyCredentialEntry(
            ConnectSecureJsonReader.GetInt(data, "id") ?? credentialId,
            ConnectSecureJsonReader.GetString(data, "name"),
            ConnectSecureJsonReader.GetString(data, "credential_type", "credentialType"),
            ConnectSecureJsonReader.GetString(data, "os_type", "osType"),
            ConnectSecureJsonReader.GetString(data, "address_type", "addressType"),
            ConnectSecureJsonReader.GetString(data, "address"),
            ConnectSecureJsonReader.GetBool(data, "is_valid", "isValid"),
            ConnectSecureJsonReader.GetString(data, "failure_reason", "failureReason"),
            ConnectSecureParamsHelper.SummarizeParams(paramsEl),
            mappedAgents);

        return (entry, paramsJson);
    }

    public async Task<int> CreateCredentialAsync(int companyId, CredentialSaveRequest request, CancellationToken ct = default)
    {
        var data = BuildCredentialPayload(companyId, request, existingParamsJson: null);
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/credentials",
            body: new Dictionary<string, object?> { ["data"] = data },
            ct: ct);

        return ExtractCreatedId(response) ?? throw new InvalidOperationException("ConnectSecure did not return a credential id.");
    }

    public async Task UpdateCredentialAsync(int credentialId, int companyId, CredentialSaveRequest request, CancellationToken ct = default)
    {
        var (_, existingParamsJson) = await GetCredentialAsync(credentialId, companyId, ct);
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            $"/r/company/credentials/{credentialId}",
            ct: ct);

        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Credential {credentialId} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        ApplyCredentialFields(dataNode, companyId, request, existingParamsJson);

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/company/credentials",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = credentialId },
            ct: ct);
    }

    public Task DeleteCredentialAsync(int credentialId, CancellationToken ct = default) =>
        client.InvokeAuthenticatedAsync(HttpMethod.Delete, $"/d/company/credentials/{credentialId}", ct: ct);

    public async Task<IReadOnlyList<CredentialMappingRecord>> ListCredentialMappingsAsync(int companyId, CancellationToken ct = default)
    {
        var rows = await FetchArrayAsync("/r/company/agent_credentials_mapping", CompanyQuery(companyId), ct);
        return rows
            .Select(row => new CredentialMappingRecord(
                ConnectSecureJsonReader.GetInt(row, "id") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "credentials_id", "credentialsId") ?? 0))
            .Where(row => row.MappingId > 0 && row.AgentId > 0 && row.CredentialsId > 0)
            .ToList();
    }

    public async Task SetProbeCredentialMappingsAsync(
        int companyId,
        int agentId,
        IReadOnlyCollection<int> credentialIds,
        CancellationToken ct = default)
    {
        var existing = (await ListCredentialMappingsAsync(companyId, ct))
            .Where(m => m.AgentId == agentId)
            .ToList();

        var desired = credentialIds.Where(id => id > 0).ToHashSet();
        var current = existing.Select(m => m.CredentialsId).ToHashSet();

        foreach (var mapping in existing.Where(m => !desired.Contains(m.CredentialsId)))
            await DeleteCredentialMappingAsync(mapping.MappingId, ct);

        foreach (var credentialId in desired.Where(id => !current.Contains(id)))
            await CreateCredentialMappingAsync(companyId, agentId, credentialId, ct);
    }

    public async Task<int> CreateCredentialMappingAsync(int companyId, int agentId, int credentialId, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/company/agent_credentials_mapping",
            body: new Dictionary<string, object?>
            {
                ["data"] = new Dictionary<string, object?>
                {
                    ["company_id"] = companyId,
                    ["agent_id"] = agentId,
                    ["credentials_id"] = credentialId
                }
            },
            ct: ct);

        return ExtractCreatedId(response) ?? throw new InvalidOperationException("ConnectSecure did not return a credential mapping id.");
    }

    public Task DeleteCredentialMappingAsync(int mappingId, CancellationToken ct = default) =>
        client.InvokeAuthenticatedAsync(HttpMethod.Delete, $"/d/company/agent_credentials_mapping/{mappingId}", ct: ct);

    public async Task SetProbeDiscoveryMappingsAsync(
        int companyId,
        int agentId,
        IReadOnlyCollection<int> discoverySettingIds,
        CancellationToken ct = default)
    {
        var rows = await FetchArrayAsync("/r/company/agent_discoverysettings_mapping", CompanyQuery(companyId), ct);
        var existing = rows
            .Select(row => new DiscoveryMappingRecord(
                ConnectSecureJsonReader.GetInt(row, "id") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "discovery_settings_id", "discoverysettings_id") ?? 0))
            .Where(row => row.MappingId > 0 && row.AgentId == agentId)
            .ToList();

        var desired = discoverySettingIds.Where(id => id > 0).ToHashSet();
        var current = existing.Select(m => m.DiscoverySettingId).ToHashSet();

        foreach (var mapping in existing.Where(m => !desired.Contains(m.DiscoverySettingId)))
        {
            await client.InvokeAuthenticatedAsync(
                HttpMethod.Delete,
                $"/d/company/agent_discoverysettings_mapping/{mapping.MappingId}",
                ct: ct);
        }

        foreach (var discoverySettingId in desired.Where(id => !current.Contains(id)))
        {
            await client.InvokeAuthenticatedAsync(
                HttpMethod.Post,
                "/w/company/agent_discoverysettings_mapping",
                body: new Dictionary<string, object?>
                {
                    ["data"] = new Dictionary<string, object?>
                    {
                        ["company_id"] = companyId,
                        ["agent_id"] = agentId,
                        ["discovery_settings_id"] = discoverySettingId
                    }
                },
                ct: ct);
        }
    }

    private static void ApplyCredentialFields(JsonObject target, int companyId, CredentialSaveRequest request, string? existingParamsJson)
    {
        var merged = BuildMergedParams(existingParamsJson, request.ParamsJson, request.MergeExistingSecrets);
        target["company_id"] = companyId;
        target["name"] = request.Name.Trim();
        target["credential_type"] = request.CredentialType.Trim();
        target["os_type"] = request.OsType ?? "";
        target["address_type"] = request.AddressType ?? "";
        target["address"] = request.Address ?? "";
        target["is_excluded"] = false;
        target["params"] = merged;
    }

    private static JsonObject BuildMergedParams(string? existingParamsJson, string incomingParamsJson, bool mergeExistingSecrets)
    {
        if (!string.IsNullOrWhiteSpace(existingParamsJson))
        {
            var existing = ConnectSecureParamsHelper.ParseParamsObject(existingParamsJson);
            var incoming = ConnectSecureParamsHelper.ParseParamsObject(incomingParamsJson);
            return ConnectSecureParamsHelper.MergeParamsJson(existing, incoming, mergeExistingSecrets);
        }

        return ConnectSecureParamsHelper.ParseParamsObject(incomingParamsJson);
    }

    private static Dictionary<string, object?> BuildCredentialPayload(
        int companyId,
        CredentialSaveRequest request,
        string? existingParamsJson)
    {
        return new Dictionary<string, object?>
        {
            ["company_id"] = companyId,
            ["name"] = request.Name.Trim(),
            ["credential_type"] = request.CredentialType.Trim(),
            ["os_type"] = request.OsType ?? "",
            ["address_type"] = request.AddressType ?? "",
            ["address"] = request.Address ?? "",
            ["is_excluded"] = false,
            ["params"] = BuildMergedParams(existingParamsJson, request.ParamsJson, request.MergeExistingSecrets)
        };
    }

    private async Task<List<JsonElement>> FetchArrayAsync(string endpoint, IReadOnlyDictionary<string, string> query, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, endpoint, query, ct: ct);
        return ConnectSecureJsonReader.ExtractDataArray(response);
    }

    private static Dictionary<string, string> CompanyQuery(int companyId) =>
        ConnectSecureCompanyReviewService.CompanyQuery(companyId, limit: 5000);

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

    public sealed record CredentialMappingRecord(int MappingId, int AgentId, int CredentialsId);

    private sealed record DiscoveryMappingRecord(int MappingId, int AgentId, int DiscoverySettingId);
}
