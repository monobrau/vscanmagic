using System.Collections.Concurrent;
using System.Text.Json;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureCacheService
{
    private readonly ConcurrentDictionary<string, CacheEntry> _entries = new(StringComparer.Ordinal);

    public static readonly TimeSpan RemediationPlanTtl = TimeSpan.FromMinutes(10);
    public static readonly TimeSpan AgentsTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan CompaniesTtl = TimeSpan.FromMinutes(30);
    public static readonly TimeSpan PatchHostsTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan CompanyReviewTtl = TimeSpan.FromMinutes(5);
    public static readonly TimeSpan AgentDetailTtl = TimeSpan.FromMinutes(5);

    public bool TryGet<T>(string key, out T value) where T : class =>
        TryGetInternal(key, out value);

    public void Set<T>(string key, T value, TimeSpan ttl) where T : class =>
        SetInternal(key, value, ttl);

    public bool TryGetRemediationPlan(int companyId, out List<JsonElement> rows) =>
        TryGet(RemediationPlanKey(companyId), out rows);

    public void SetRemediationPlan(int companyId, List<JsonElement> rows) =>
        Set(RemediationPlanKey(companyId), rows, RemediationPlanTtl);

    public bool TryGetPatchHosts(int companyId, int solutionId, out List<PatchAssetDetail> details) =>
        TryGet(PatchHostsKey(companyId, solutionId), out details);

    public void SetPatchHosts(int companyId, int solutionId, List<PatchAssetDetail> details) =>
        Set(PatchHostsKey(companyId, solutionId), details, PatchHostsTtl);

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

        _entries.TryRemove(RemediationPlanKey(companyId), out _);
        _entries.TryRemove(CompanyReviewKey(companyId), out _);
    }

    public void InvalidatePatchHosts(int companyId)
    {
        var prefix = $"patch_hosts:{companyId}:";
        foreach (var key in _entries.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)))
            _entries.TryRemove(key, out _);
    }

    public void InvalidateRemediationPlan(int companyId) =>
        _entries.TryRemove(RemediationPlanKey(companyId), out _);

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

    private static string RemediationPlanKey(int companyId) => $"remediation_plan:{companyId}";

    private static string PatchHostsKey(int companyId, int solutionId) => $"patch_hosts:{companyId}:{solutionId}";

    private static string AgentDetailKey(int agentId) => $"agent_detail:{agentId}";

    private static string CompanyReviewKey(int companyId) => $"company_review:{companyId}";

    private sealed record CacheEntry(object Value, DateTimeOffset ExpiresAt);
}
