using VScanMagic.Core.Configuration;

namespace VScanMagic.Core.Risk;

public static class ProductConsolidator
{
    public static string GetConsolidatedProduct(string productName, VScanMagicOptions options)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return productName;

        var normalized = productName.Trim();
        var groupKey = GetTimeEstimateGroupKey(normalized);
        if (groupKey != normalized)
            return groupKey;

        foreach (var (consolidated, patterns) in options.WindowsConsolidation)
        {
            foreach (var pattern in patterns)
            {
                if (normalized.Equals(pattern, StringComparison.OrdinalIgnoreCase) ||
                    normalized.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return consolidated;
            }
        }

        var versionMatch = System.Text.RegularExpressions.Regex.Match(normalized, @"^(.+?)\s+[\d\.]+$");
        if (versionMatch.Success)
        {
            var baseProduct = versionMatch.Groups[1].Value;
            foreach (var (consolidated, patterns) in options.WindowsConsolidation)
            {
                foreach (var pattern in patterns)
                {
                    if (baseProduct.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                        return consolidated;
                }
            }
        }

        return productName;
    }

    public static string GetTimeEstimateGroupKey(string productName)
    {
        if (productName.Contains("Visual C++", StringComparison.OrdinalIgnoreCase) ||
            productName.Contains("Microsoft Visual C++", StringComparison.OrdinalIgnoreCase))
            return "Microsoft Visual C++ (all versions)";

        if (productName.Contains(".NET", StringComparison.OrdinalIgnoreCase) &&
            (productName.Contains("Runtime", StringComparison.OrdinalIgnoreCase) ||
             productName.Contains("Framework", StringComparison.OrdinalIgnoreCase) ||
             productName.Contains("SDK", StringComparison.OrdinalIgnoreCase)))
            return "Microsoft .NET (all versions)";

        return productName;
    }

    /// <summary>
    /// Trims noisy patch versions for display (e.g. "MongoDB 3.4.24" → "MongoDB 3.4").
    /// </summary>
    public static string GetProductMajorVersion(string productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return productName;

        var p = productName.Trim();
        var majorMinorPatch = System.Text.RegularExpressions.Regex.Match(p, @"^(.+?)\s+(\d+\.\d+)(\.\d+)+$");
        if (majorMinorPatch.Success)
            return $"{majorMinorPatch.Groups[1].Value.Trim()} {majorMinorPatch.Groups[2].Value}";

        var trailingVersion = System.Text.RegularExpressions.Regex.Match(p, @"^(.+?)\s+\d+\.\d+(\.\d+)*$");
        if (trailingVersion.Success)
            return trailingVersion.Groups[1].Value.Trim();

        return p;
    }

    public static string FormatDisplayName(string productName, VScanMagicOptions? options = null) =>
        ProductNameNormalizer.FormatDisplayName(productName, options);
}
