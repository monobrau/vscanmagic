using System.Text.RegularExpressions;

namespace VScanMagic.Core.Risk;

public static partial class CveReferenceHelper
{
    public static IReadOnlyList<string> SplitCveIds(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return [];

        return CveIdPattern().Matches(text)
            .Select(m => m.Value.ToUpperInvariant())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static string GetNvdDetailUrl(string cveId) =>
        $"https://nvd.nist.gov/vuln/detail/{cveId.ToUpperInvariant()}";

    public static string FormatReferenceLinks(string? cveIdsText, string? separator = null)
    {
        var cves = SplitCveIds(cveIdsText);
        if (cves.Count == 0)
            return "";

        separator ??= "; ";
        return string.Join(separator, cves);
    }

    public static string MergeCveIds(IEnumerable<string?>? sources)
    {
        var merged = (sources ?? [])
            .SelectMany(value => SplitCveIds(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return merged.Count == 0 ? "" : string.Join("; ", merged);
    }

    public static string MergeCveIds(params string?[] values) => MergeCveIds(values.AsEnumerable());

    public static string NormalizeFindingCveIds(string? cveIds, string? product) =>
        MergeCveIds(cveIds, IsCveOnlyProduct(product) ? product : null);

    public static bool IsHighSeverityCveOnlyProduct(string? productName, int critical, int high) =>
        IsCveOnlyProduct(productName) && (critical > 0 || high > 0);

    /// <summary>True when the product/name field is only a CVE identifier (not a real software product).</summary>
    public static bool IsCveOnlyProduct(string? productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return false;

        return CveOnlyProductPattern().IsMatch(productName.Trim());
    }

    [GeneratedRegex(@"^CVE-\d{4}-\d+$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex CveOnlyProductPattern();

    [GeneratedRegex(@"CVE-\d{4}-\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
    private static partial Regex CveIdPattern();
}
