using System.Text.Json;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureProbeConfigurationService(
    ConnectSecureClient client,
    ConnectSecureCompanyReviewService reviewService,
    ConnectSecureCompanyCredentialsService credentialsService,
    ConnectSecureIntegrationService integrationService)
{
    public async Task<ProbeConfigurationData> GetConfigurationAsync(int companyId, string companyName = "", CancellationToken ct = default)
    {
        var result = new ProbeConfigurationData { CompanyId = companyId, CompanyName = companyName };
        if (companyId <= 0)
            return result;

        var review = await reviewService.GetReviewDataAsync(companyId, companyName, ct);
        result.InternalSubnets.AddRange(review.InternalSubnets);

        var credentials = await credentialsService.ListCredentialsAsync(companyId, ct);
        result.Credentials.AddRange(credentials);

        var credentialMappings = await credentialsService.ListCredentialMappingsAsync(companyId, ct);
        var discoveryMappings = await FetchDiscoveryMappingsAsync(companyId, ct);

        foreach (var probe in review.ProbeAgentsNmapInfo.Where(p => p.AgentId is > 0))
        {
            var agentId = probe.AgentId!.Value;
            result.Probes.Add(new ProbeAgentConfiguration(
                agentId,
                probe.HostName,
                probe.Ip,
                credentialMappings.Where(m => m.AgentId == agentId).Select(m => m.CredentialsId).Distinct().ToList(),
                discoveryMappings.Where(m => m.AgentId == agentId).Select(m => m.DiscoverySettingId).Distinct().ToList()));
        }

        var integrations = await integrationService.ListIntegrationCredentialsAsync(companyId, ct);
        result.Integrations.AddRange(integrations);

        var mappings = await integrationService.ListCompanyMappingsAsync(companyId, ct);
        result.IntegrationMappings.AddRange(mappings);

        result.KnownCredentialTypes = CredentialTypeCatalog.MergeKnownTypes(credentials.Select(c => c.CredentialType));
        result.KnownIntegrationNames = integrations
            .Select(i => i.IntegrationName)
            .Concat(mappings.Select(m => m.IntegrationName))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return result;
    }

    public Task SetProbeCredentialMappingsAsync(int companyId, int agentId, IReadOnlyCollection<int> credentialIds, CancellationToken ct = default) =>
        credentialsService.SetProbeCredentialMappingsAsync(companyId, agentId, credentialIds, ct);

    public Task SetProbeDiscoveryMappingsAsync(int companyId, int agentId, IReadOnlyCollection<int> discoverySettingIds, CancellationToken ct = default) =>
        credentialsService.SetProbeDiscoveryMappingsAsync(companyId, agentId, discoverySettingIds, ct);

    private async Task<IReadOnlyList<DiscoveryMappingRecord>> FetchDiscoveryMappingsAsync(int companyId, CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            "/r/company/agent_discoverysettings_mapping",
            ConnectSecureCompanyReviewService.CompanyQuery(companyId, limit: 5000),
            ct: ct);

        return ConnectSecureJsonReader.ExtractDataArray(response)
            .Select(row => new DiscoveryMappingRecord(
                ConnectSecureJsonReader.GetInt(row, "id") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "agent_id", "agentId") ?? 0,
                ConnectSecureJsonReader.GetInt(row, "discovery_settings_id", "discoverysettings_id") ?? 0))
            .Where(row => row.MappingId > 0 && row.AgentId > 0 && row.DiscoverySettingId > 0)
            .ToList();
    }

    private sealed record DiscoveryMappingRecord(int MappingId, int AgentId, int DiscoverySettingId);
}
