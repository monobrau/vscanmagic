using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public sealed record TimeEstimateGroupRow(
    string GroupKey,
    ReviewFinding SourceFinding,
    int HostCount);

public static class ReviewFindingTimeEstimate
{
    public static string GetGroupKey(ReviewFinding finding) =>
        ProductConsolidator.GetTimeEstimateGroupKey(finding.Product);

    public static IReadOnlyList<TimeEstimateGroupRow> GetExportGroups(ReviewSession session)
    {
        var byKey = new Dictionary<string, (ReviewFinding Source, HashSet<string> Hosts)>(StringComparer.OrdinalIgnoreCase);

        foreach (var finding in ReviewSessionRanker.GetExportFindings(session))
        {
            var key = GetGroupKey(finding);
            if (!byKey.TryGetValue(key, out var entry))
            {
                entry = (finding, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                byKey[key] = entry;
            }
            else if (finding.OriginalRank < entry.Source.OriginalRank)
            {
                byKey[key] = (finding, entry.Hosts);
                entry = byKey[key];
            }

            foreach (var system in finding.AffectedSystems ?? [])
            {
                if (!system.ExcludedFromExport)
                    entry.Hosts.Add($"{system.HostName}|{system.Ip}");
            }
        }

        return byKey
            .Select(pair => new TimeEstimateGroupRow(pair.Key, pair.Value.Source, pair.Value.Hosts.Count))
            .OrderBy(row => row.SourceFinding.OriginalRank)
            .ToList();
    }

    public static void ApplyToGroup(ReviewSession session, ReviewFinding source)
    {
        var key = GetGroupKey(source);
        foreach (var finding in session.Findings)
        {
            if (!string.Equals(GetGroupKey(finding), key, StringComparison.Ordinal))
                continue;

            finding.TimeEstimateHours = source.TimeEstimateHours;
            finding.AfterHours = source.AfterHours;
            finding.ThirdParty = source.ThirdParty;
            finding.TicketGenerated = source.TicketGenerated;
        }
    }

    public static void EnsureDefaults(ReviewSession session)
    {
        foreach (var finding in session.Findings)
        {
            if (finding.TimeEstimateInitialized)
                continue;

            finding.ThirdParty = FirstPartyVendorHelper.IsThirdPartyByDefault(GetGroupKey(finding), session.IsRmitPlus);
            finding.TimeEstimateInitialized = true;
        }
    }
}
