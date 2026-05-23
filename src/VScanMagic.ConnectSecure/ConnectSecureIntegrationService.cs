using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureIntegrationService(ConnectSecureClient client)
{
    public async Task<IReadOnlyList<IntegrationCredentialEntry>> ListIntegrationCredentialsAsync(int companyId, CancellationToken ct = default)
    {
        var rows = await FetchArrayAsync("/r/integration/integration_credentials", CompanyQuery(companyId), ct);
        return rows
            .Select(row =>
            {
                row.TryGetProperty("params", out var paramsEl);
                return new IntegrationCredentialEntry(
                    ConnectSecureJsonReader.GetInt(row, "id") ?? 0,
                    ConnectSecureJsonReader.GetString(row, "name"),
                    ConnectSecureJsonReader.GetString(row, "integration_name", "integrationName"),
                    ConnectSecureJsonReader.GetString(row, "ticket_url", "ticketUrl"),
                    ConnectSecureParamsHelper.SummarizeParams(paramsEl));
            })
            .Where(row => row.Id > 0)
            .OrderBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<(IntegrationCredentialEntry Entry, string ParamsJson)> GetIntegrationCredentialAsync(int id, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, $"/r/integration/integration_credentials/{id}", ct: ct);
        if (!response.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Integration credential {id} was not found.");

        data.TryGetProperty("params", out var paramsEl);
        var entry = new IntegrationCredentialEntry(
            ConnectSecureJsonReader.GetInt(data, "id") ?? id,
            ConnectSecureJsonReader.GetString(data, "name"),
            ConnectSecureJsonReader.GetString(data, "integration_name", "integrationName"),
            ConnectSecureJsonReader.GetString(data, "ticket_url", "ticketUrl"),
            ConnectSecureParamsHelper.SummarizeParams(paramsEl));

        return (entry, paramsEl.ValueKind == JsonValueKind.Object ? paramsEl.GetRawText() : "{}");
    }

    public async Task<int> CreateIntegrationCredentialAsync(int companyId, IntegrationCredentialSaveRequest request, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/integration/integration_credentials",
            body: new Dictionary<string, object?> { ["data"] = BuildIntegrationCredentialPayload(companyId, request, null) },
            ct: ct);

        return ExtractCreatedId(response) ?? throw new InvalidOperationException("ConnectSecure did not return an integration credential id.");
    }

    public async Task UpdateIntegrationCredentialAsync(int id, int companyId, IntegrationCredentialSaveRequest request, CancellationToken ct = default)
    {
        var (_, existingParamsJson) = await GetIntegrationCredentialAsync(id, ct);
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, $"/r/integration/integration_credentials/{id}", ct: ct);
        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Integration credential {id} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        var payload = BuildIntegrationCredentialPayload(companyId, request, existingParamsJson);
        foreach (var property in payload)
            dataNode[property.Key] = JsonNode.Parse(JsonSerializer.Serialize(property.Value));

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/integration/integration_credentials",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = id },
            ct: ct);
    }

    public Task DeleteIntegrationCredentialAsync(int id, CancellationToken ct = default) =>
        client.InvokeAuthenticatedAsync(HttpMethod.Delete, $"/d/integration/integration_credentials/{id}", ct: ct);

    public async Task<IReadOnlyList<CompanyIntegrationMappingEntry>> ListCompanyMappingsAsync(int companyId, CancellationToken ct = default)
    {
        var rows = await FetchArrayAsync("/r/integration/company_mappings", CompanyQuery(companyId), ct);
        return rows
            .Select(row => new CompanyIntegrationMappingEntry(
                ConnectSecureJsonReader.GetInt(row, "id") ?? 0,
                ConnectSecureJsonReader.GetString(row, "integration_name", "integrationName"),
                ConnectSecureJsonReader.GetString(row, "source_company_name", "sourceCompanyName"),
                ConnectSecureJsonReader.GetString(row, "dest_company_name", "destCompanyName"),
                ConnectSecureJsonReader.GetString(row, "site_name", "siteName"),
                ConnectSecureJsonReader.GetInt(row, "credential_id", "credentialId")))
            .Where(row => row.Id > 0)
            .OrderBy(row => row.IntegrationName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<int> CreateCompanyMappingAsync(int companyId, IntegrationMappingSaveRequest request, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/integration/company_mappings",
            body: new Dictionary<string, object?> { ["data"] = BuildCompanyMappingPayload(companyId, request, null) },
            ct: ct);

        return ExtractCreatedId(response) ?? throw new InvalidOperationException("ConnectSecure did not return an integration mapping id.");
    }

    public async Task UpdateCompanyMappingAsync(int id, int companyId, IntegrationMappingSaveRequest request, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, $"/r/integration/company_mappings/{id}", ct: ct);
        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Integration mapping {id} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        var payload = BuildCompanyMappingPayload(companyId, request, dataEl.GetRawText());
        foreach (var property in payload)
            dataNode[property.Key] = JsonNode.Parse(JsonSerializer.Serialize(property.Value));

        await client.InvokeAuthenticatedAsync(
            HttpMethod.Patch,
            "/w/integration/company_mappings",
            body: new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = id },
            ct: ct);
    }

    public Task DeleteCompanyMappingAsync(int id, CancellationToken ct = default) =>
        client.InvokeAuthenticatedAsync(HttpMethod.Delete, $"/d/integration/company_mappings/{id}", ct: ct);

    private static Dictionary<string, object?> BuildIntegrationCredentialPayload(
        int companyId,
        IntegrationCredentialSaveRequest request,
        string? existingParamsJson)
    {
        var existing = ConnectSecureParamsHelper.ParseParamsObject(existingParamsJson);
        var incoming = ConnectSecureParamsHelper.ParseParamsObject(request.ParamsJson);
        var merged = ConnectSecureParamsHelper.MergeParamsJson(existing, incoming, request.MergeExistingSecrets);

        return new Dictionary<string, object?>
        {
            ["company_id"] = companyId,
            ["name"] = request.Name.Trim(),
            ["integration_name"] = request.IntegrationName.Trim(),
            ["ticket_url"] = request.TicketUrl ?? "",
            ["params"] = merged
        };
    }

    private static Dictionary<string, object?> BuildCompanyMappingPayload(
        int companyId,
        IntegrationMappingSaveRequest request,
        string? existingJson)
    {
        var payload = new Dictionary<string, object?>
        {
            ["company_id"] = companyId,
            ["integration_name"] = request.IntegrationName.Trim(),
            ["source_company_name"] = request.SourceCompanyName ?? "",
            ["dest_company_name"] = request.DestCompanyName ?? "",
            ["dest_company_id"] = request.DestCompanyId ?? "",
            ["site_name"] = request.SiteName ?? "",
            ["site_id"] = request.SiteId ?? "",
            ["credential_id"] = request.CredentialId
        };

        if (!string.IsNullOrWhiteSpace(request.ParamsJson) && request.ParamsJson.Trim() != "{}")
        {
            var extra = ConnectSecureParamsHelper.ParseParamsObject(request.ParamsJson);
            foreach (var property in extra)
                payload[property.Key] = property.Value?.DeepClone();
        }

        return payload;
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
}
