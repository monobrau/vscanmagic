using System.Collections.Concurrent;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureCacheService
{
    private readonly ConcurrentDictionary<string, CacheEntry> _entries = new(StringComparer.Ordinal);

    public static readonly TimeSpan RemediationDatasetTtl = TimeSpan.FromMinutes(10);
    public static readonly TimeSpan AgentsTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan CompaniesTtl = TimeSpan.FromMinutes(30);
    public static readonly TimeSpan PatchHostsTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan RemediationAssetDetailsTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan CompanyReviewTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan AgentDetailTtl = TimeSpan.FromMinutes(5);

    public bool TryGet<T>(string key, out T value) where T : class =>
        TryGetInternal(key, out value);

    public void Set<T>(string key, T value, TimeSpan ttl) where T : class =>
        SetInternal(key, value, ttl);

    public bool TryGetRemediationDataset(int companyId, out RemediationDataset dataset) =>
        TryGet(RemediationDatasetKey(companyId), out dataset);

    public void SetRemediationDataset(int companyId, RemediationDataset dataset) =>
        Set(RemediationDatasetKey(companyId), dataset, RemediationDatasetTtl);

    public bool TryGetPatchHosts(int companyId, int solutionId, out List<PatchAssetDetail> details) =>
        TryGet(PatchHostsKey(companyId, solutionId), out details);

    public void SetPatchHosts(int companyId, int solutionId, List<PatchAssetDetail> details) =>
        Set(PatchHostsKey(companyId, solutionId), details, PatchHostsTtl);

    public bool TryGetRemediationAssetDetails(int companyId, out List<PatchAssetDetail> details) =>
        TryGet(RemediationAssetDetailsKey(companyId), out details);

    public void SetRemediationAssetDetails(int companyId, List<PatchAssetDetail> details) =>
        Set(RemediationAssetDetailsKey(companyId), details, RemediationAssetDetailsTtl);

    public bool TryGetProductRemediationAssetDetails(int companyId, string productName, out List<PatchAssetDetail> details) =>
        TryGet(ProductRemediationAssetDetailsKey(companyId, productName), out details);

    public void SetProductRemediationAssetDetails(int companyId, string productName, List<PatchAssetDetail> details) =>
        Set(ProductRemediationAssetDetailsKey(companyId, productName), details, RemediationAssetDetailsTtl);

    public bool TryGetPatchProductHosts(int companyId, string productName, out List<PatchAssetDetail> details) =>
        TryGet(PatchProductHostsKey(companyId, productName), out details);

    public void SetPatchProductHosts(int companyId, string productName, List<PatchAssetDetail> details) =>
        Set(PatchProductHostsKey(companyId, productName), details, PatchHostsTtl);

    public bool TryGetAgentDetail(int agentId, out PatchAssetDetail detail) =>
        TryGet(AgentDetailKey(agentId), out detail);

    public void SetAgentDetail(int agentId, PatchAssetDetail detail) =>
        Set(AgentDetailKey(agentId), detail, AgentDetailTtl);

    public bool TryGetCompanyReview(int companyId, out CompanyReviewData data) =>
        TryGet(CompanyReviewKey(companyId), out data);

    public void SetCompanyReview(int companyId, CompanyReviewData data) =>
        Set(CompanyReviewKey(companyId), data, CompanyReviewTtl);

    public void InvalidateCompany(int companyId)
    {
        var prefix = $"company:{companyId}:";
        foreach (var key in _entries.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)))
            _entries.TryRemove(key, out _);

        _entries.TryRemove(RemediationDatasetKey(companyId), out _);
        _entries.TryRemove(CompanyReviewKey(companyId), out _);
        InvalidatePatchHosts(companyId);
    }

    public void InvalidatePatchHosts(int companyId)
    {
        foreach (var prefix in new[]
                 {
                     $"patch_hosts:{companyId}:",
                     $"patch_product_hosts:{companyId}:",
                     $"remediation_asset_details:{companyId}:"
                 })
        {
            foreach (var key in _entries.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)))
                _entries.TryRemove(key, out _);
        }

        _entries.TryRemove(RemediationAssetDetailsKey(companyId), out _);
    }

    public void InvalidateRemediationDataset(int companyId) =>
        _entries.TryRemove(RemediationDatasetKey(companyId), out _);

    public void InvalidateCompanyReview(int companyId) =>
        _entries.TryRemove(CompanyReviewKey(companyId), out _);

    private bool TryGetInternal<T>(string key, out T value) where T : class
    {
        value = default!;
        if (!_entries.TryGetValue(key, out var entry))
            return false;

        if (DateTimeOffset.UtcNow >= entry.ExpiresAt)
        {
            _entries.TryRemove(key, out _);
            return false;
        }

        if (entry.Value is T typed)
        {
            value = typed;
            return true;
        }

        return false;
    }

    private void SetInternal<T>(string key, T value, TimeSpan ttl) where T : class =>
        _entries[key] = new CacheEntry(value, DateTimeOffset.UtcNow.Add(ttl));

    private static string RemediationDatasetKey(int companyId) => $"remediation_dataset:{companyId}";

    private static string PatchHostsKey(int companyId, int solutionId) => $"patch_hosts:{companyId}:{solutionId}";

    private static string RemediationAssetDetailsKey(int companyId) => $"remediation_asset_details:{companyId}";

    private static string ProductRemediationAssetDetailsKey(int companyId, string productName) =>
        $"remediation_asset_details:{companyId}:{NormalizeProductKey(productName)}";

    private static string PatchProductHostsKey(int companyId, string productName) =>
        $"patch_product_hosts:{companyId}:{NormalizeProductKey(productName)}";

    private static string NormalizeProductKey(string productName) =>
        productName.Trim().ToLowerInvariant();

    private static string AgentDetailKey(int agentId) => $"agent_detail:{agentId}";

    private static string CompanyReviewKey(int companyId) => $"company_review:{companyId}";

    private sealed record CacheEntry(object Value, DateTimeOffset ExpiresAt);
}
