namespace VScanMagic.ConnectSecure;

/// <summary>
/// Maps ConnectSecure asset hostnames to logged-in users with short-name (NetBIOS) fallbacks.
/// </summary>
internal static class HostnameUsernameMatcher
{
    public static void RegisterHost(Dictionary<string, string> assetToUser, string hostname, string username)
    {
        if (string.IsNullOrWhiteSpace(hostname) || string.IsNullOrWhiteSpace(username))
            return;

        var user = username.Trim();
        foreach (var key in HostnameKeys(hostname))
            assetToUser[key] = user;
    }

    public static bool TryResolveUser(
        IReadOnlyDictionary<string, string> assetToUser,
        string hostname,
        out string username)
    {
        username = "";
        if (string.IsNullOrWhiteSpace(hostname))
            return false;

        foreach (var key in HostnameKeys(hostname))
        {
            if (assetToUser.TryGetValue(key, out var user) && !string.IsNullOrWhiteSpace(user))
            {
                username = user;
                return true;
            }
        }

        return false;
    }

    public static IEnumerable<string> HostnameKeys(string hostname)
    {
        var normalized = hostname.Trim().ToLowerInvariant();
        yield return normalized;

        var shortName = ShortHostname(normalized);
        if (!string.Equals(shortName, normalized, StringComparison.Ordinal))
            yield return shortName;
    }

    private static string ShortHostname(string normalized)
    {
        var dot = normalized.IndexOf('.');
        return dot > 0 ? normalized[..dot] : normalized;
    }
}
