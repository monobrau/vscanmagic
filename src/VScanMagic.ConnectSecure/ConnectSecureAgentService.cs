using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureAgentService(ConnectSecureClient client)
{
    public async Task UpdateNmapInterfaceAsync(int agentId, string nmapInterface, CancellationToken ct = default)
    {
        var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, $"/r/company/agents/{agentId}", ct: ct);
        if (!response.TryGetProperty("data", out var dataEl) || dataEl.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException($"Agent {agentId} was not found.");

        var dataNode = JsonNode.Parse(dataEl.GetRawText())!.AsObject();
        dataNode["nmap_interface"] = nmapInterface;

        var attempts = new (HttpMethod Method, string Endpoint, Dictionary<string, object?> Body)[]
        {
            (HttpMethod.Patch, "/w/company/agents", new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = agentId }),
            (HttpMethod.Patch, "/w/company/agent_discovery_credentials", new Dictionary<string, object?> { ["data"] = dataNode, ["id"] = agentId }),
            (HttpMethod.Post, "/w/company/agents", new Dictionary<string, object?> { ["data"] = dataNode }),
        };

        HttpRequestException? lastError = null;
        foreach (var (method, endpoint, body) in attempts)
        {
            try
            {
                await client.InvokeAuthenticatedAsync(method, endpoint, body: body, ct: ct);
                return;
            }
            catch (HttpRequestException ex) when (ex.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.MethodNotAllowed)
            {
                lastError = ex;
            }
            catch (HttpRequestException ex)
            {
                lastError = ex;
                if (ex.Message.Contains("unknown field", StringComparison.OrdinalIgnoreCase))
                    continue;
                throw;
            }
        }

        throw new InvalidOperationException(
            "ConnectSecure did not accept an nmap interface update. The portal may use an undocumented endpoint — capture the network request when saving in the portal and we can wire the exact path.",
            lastError);
    }
}
