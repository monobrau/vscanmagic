using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class ReviewUsernameLookup
{
    /// <summary>
    /// Fills empty <see cref="ReviewAffectedSystem.Username"/> values from a hostname lookup map.
    /// Returns the number of hosts updated.
    /// </summary>
    public static int ApplyEmptyUsernames(ReviewSession session, IReadOnlyDictionary<string, string> lookupByHostname)
    {
        if (lookupByHostname.Count == 0)
            return 0;

        var updated = 0;
        foreach (var finding in session.Findings)
        {
            foreach (var system in finding.AffectedSystems ?? [])
            {
                if (!string.IsNullOrWhiteSpace(system.Username))
                    continue;

                var host = system.HostName?.Trim();
                if (string.IsNullOrWhiteSpace(host))
                    continue;

                if (!lookupByHostname.TryGetValue(host, out var user) || string.IsNullOrWhiteSpace(user))
                    continue;

                system.Username = user.Trim();
                updated++;
            }
        }

        return updated;
    }
}
