using System.Collections.Concurrent;
using VScanMagic.ConnectSecure;

namespace VScanMagic.Web.Services;

public sealed class CompanyListService(ConnectSecureClient client)
{
    private readonly SemaphoreSlim _loadLock = new(1, 1);
    private IReadOnlyList<CompanyInfo>? _cached;
    private DateTimeOffset _expiresAt = DateTimeOffset.MinValue;

    public async Task<IReadOnlyList<CompanyInfo>> GetCompaniesAsync(CancellationToken ct = default)
    {
        if (_cached is not null && DateTimeOffset.UtcNow < _expiresAt)
            return _cached;

        await _loadLock.WaitAsync(ct);
        try
        {
            if (_cached is not null && DateTimeOffset.UtcNow < _expiresAt)
                return _cached;

            _cached = await client.GetCompaniesAsync(ct);
            _expiresAt = DateTimeOffset.UtcNow.Add(ConnectSecureCacheService.CompaniesTtl);
            return _cached;
        }
        finally
        {
            _loadLock.Release();
        }
    }

    public void Invalidate() => _expiresAt = DateTimeOffset.MinValue;
}
