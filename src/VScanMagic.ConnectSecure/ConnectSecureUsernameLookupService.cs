using System.Text.Json;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureUsernameLookupService(ConnectSecureClient client)
{
    private const int PageSize = 500;

    public bool IsConfigured => client.IsConfigured;

    /// <summary>
    /// Returns a map from requested hostname to logged_in_user from ConnectSecure asset_view.
    /// Only hostnames with a match are included; values may be empty when not found.
    /// </summary>
    public async Task<IReadOnlyDictionary<string, string>> GetUsernamesByHostnameAsync(
        int companyId,
        IEnumerable<string> hostnames,
        CancellationToken ct = default)
    {
        var requested = hostnames
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Select(h => h.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var result = requested.ToDictionary(h => h, _ => "", StringComparer.OrdinalIgnoreCase);
        if (companyId <= 0 || requested.Count == 0)
            return result;

        if (!client.IsConfigured)
            throw new InvalidOperationException("ConnectSecure is not configured. Add credentials in Settings.");

        var assetToUser = await LoadAssetUsernameIndexAsync(companyId, ct);

        foreach (var host in requested)
        {
            if (HostnameUsernameMatcher.TryResolveUser(assetToUser, host, out var user))
                result[host] = user;
        }

        return result;
    }

    private async Task<Dictionary<string, string>> LoadAssetUsernameIndexAsync(int companyId, CancellationToken ct)
    {
        var assetToUser = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var skip = 0;

        while (true)
        {
            var query = ConnectSecureCompanyReviewService.CompanyQuery(companyId, limit: PageSize, skip: skip, orderBy: "host_name asc");
            var response = await client.InvokeAuthenticatedAsync(HttpMethod.Get, "/r/asset/asset_view", query, ct: ct);
            var rows = ConnectSecureJsonReader.ExtractDataArray(response);
            if (rows.Count == 0)
                break;

            foreach (var row in rows)
            {
                var host = ConnectSecureJsonReader.GetString(
                    row, "host_name", "hostname", "name", "Host Name");
                var user = ConnectSecureJsonReader.GetString(
                    row, "logged_in_user", "logged_in_user_name", "Logged In User");
                HostnameUsernameMatcher.RegisterHost(assetToUser, host, user);

                var nameAlias = ConnectSecureJsonReader.GetString(row, "name");
                if (!string.IsNullOrWhiteSpace(nameAlias) &&
                    !string.Equals(nameAlias, host, StringComparison.OrdinalIgnoreCase))
                    HostnameUsernameMatcher.RegisterHost(assetToUser, nameAlias, user);
            }

            skip += rows.Count;
            if (rows.Count < PageSize)
                break;
        }

        return assetToUser;
    }
}
