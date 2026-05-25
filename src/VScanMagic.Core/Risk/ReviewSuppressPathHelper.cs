namespace VScanMagic.Core.Risk;

public static class ReviewSuppressPathHelper
{
    /// <summary>
    /// Application findings without CVE references suppress via remediation solution ids.
    /// Everything else (CVE-only product, registry, network, app+CVE) uses problem ids.
    /// </summary>
    public static bool UsesApplicationSolutionPath(string? source, string? product, string? cveIds) =>
        VulnerabilitySourceHelper.IsApplication(source) &&
        CveReferenceHelper.SplitCveIds(CveReferenceHelper.NormalizeFindingCveIds(cveIds, product)).Count == 0;
}
