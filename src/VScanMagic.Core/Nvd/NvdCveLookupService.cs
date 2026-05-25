using System.Net.Http.Headers;
using System.Text.Json;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Risk;

namespace VScanMagic.Core.Nvd;

public sealed class NvdCveLookupService
{
    private static readonly Uri BaseUri = new("https://services.nvd.nist.gov/rest/json/cves/2.0");
    private readonly HttpClient _httpClient;
    private readonly object _rateLimitLock = new();
    private DateTime _lastRequestUtc = DateTime.MinValue;

    public NvdCveLookupService()
        : this(new HttpClient())
    {
    }

    internal NvdCveLookupService(HttpClient httpClient)
    {
        _httpClient = httpClient;
        _httpClient.Timeout = TimeSpan.FromSeconds(60);
    }

    public async Task<string> GetRemediationSummaryAsync(string cveId, string? apiKey, CancellationToken ct = default)
    {
        var normalized = cveId.Trim().ToUpperInvariant();
        if (!System.Text.RegularExpressions.Regex.IsMatch(normalized, @"^CVE-\d{4}-\d+$"))
            return "";

        var cached = TryReadCache(normalized);
        if (cached is not null)
            return cached;

        await WaitForRateLimitAsync(apiKey, ct).ConfigureAwait(false);

        using var request = new HttpRequestMessage(HttpMethod.Get, $"{BaseUri}?cveId={Uri.EscapeDataString(normalized)}");
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (!string.IsNullOrWhiteSpace(apiKey))
            request.Headers.Add("apiKey", apiKey.Trim());

        try
        {
            using var response = await _httpClient.SendAsync(request, ct).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);

            var summary = "";
            if (document.RootElement.TryGetProperty("vulnerabilities", out var vulnerabilities) &&
                vulnerabilities.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in vulnerabilities.EnumerateArray())
                {
                    summary = NvdRemediationFormatter.BuildSummary(item);
                    if (!string.IsNullOrWhiteSpace(summary))
                        break;
                }
            }

            WriteCache(normalized, summary);
            return summary;
        }
        catch
        {
            WriteCache(normalized, "");
            return "";
        }
    }

    public async Task<string> GetRemediationSummaryForListAsync(
        IEnumerable<string> cveIds,
        string? apiKey,
        CancellationToken ct = default)
    {
        var chunks = new List<string>();
        foreach (var id in cveIds.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var summary = await GetRemediationSummaryAsync(id, apiKey, ct).ConfigureAwait(false);
            if (!string.IsNullOrWhiteSpace(summary))
                chunks.Add(summary);
        }

        var combined = string.Join(" | ", chunks);
        if (combined.Length > 500)
            combined = combined[..500].TrimEnd() + "…";

        return combined;
    }

    private async Task WaitForRateLimitAsync(string? apiKey, CancellationToken ct)
    {
        var minDelayMs = string.IsNullOrWhiteSpace(apiKey) ? 6000 : 600;
        int waitMs;

        lock (_rateLimitLock)
        {
            var elapsed = (DateTime.UtcNow - _lastRequestUtc).TotalMilliseconds;
            waitMs = elapsed < minDelayMs ? (int)Math.Ceiling(minDelayMs - elapsed) : 0;
        }

        if (waitMs > 0)
            await Task.Delay(waitMs, ct).ConfigureAwait(false);

        lock (_rateLimitLock)
            _lastRequestUtc = DateTime.UtcNow;
    }

    private static string? TryReadCache(string cveId)
    {
        var path = VScanMagicPaths.NvdCacheFile(cveId);
        if (!File.Exists(path))
            return null;

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(path));
            if (document.RootElement.TryGetProperty("summary", out var summaryElement))
                return summaryElement.GetString() ?? "";
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static void WriteCache(string cveId, string summary)
    {
        var path = VScanMagicPaths.NvdCacheFile(cveId);
        var payload = JsonSerializer.Serialize(new
        {
            cveId,
            fetchedAt = DateTimeOffset.UtcNow,
            summary
        });
        File.WriteAllText(path, payload);
    }
}
