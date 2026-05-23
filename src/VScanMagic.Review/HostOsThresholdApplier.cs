using VScanMagic.Core.Models;
using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class HostOsThresholdApplier
{
    public static void Apply(ReviewSession session, UserSettings settings)
    {
        foreach (var finding in session.Findings)
        {
            var systems = finding.AffectedSystems ?? [];
            if (systems.Count == 0)
                continue;

            foreach (var system in systems)
            {
                system.ExcludedFromExport = !OsHostVulnThresholdHelper.ShouldIncludeHost(
                    finding.Product, finding.Source, system.VulnCount, settings);
            }
        }
    }
}
