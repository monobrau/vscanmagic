using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

/// <summary>Review UI helpers for findings whose product name is only a CVE id.</summary>
public static class CveOnlyFindingDisplay
{
    public const string Explanation =
        "ConnectSecure reported this by CVE ID only (common for registry or network findings). " +
        "Use the NVD summary and affected hosts to identify the software, firmware, or device.";

    public static bool IsCveOnlyFinding(ReviewFinding finding) =>
        CveReferenceHelper.IsCveOnlyProduct(finding.Product);

    public static string GetListSubtitle(ReviewFinding finding)
    {
        if (!IsCveOnlyFinding(finding))
            return "";

        if (!string.IsNullOrWhiteSpace(finding.NvdEnrichment))
        {
            var description = ParseNvdEnrichment(finding.NvdEnrichment).Description;
            if (!string.IsNullOrWhiteSpace(description))
                return Truncate(description, 120);

            return "NVD references loaded — select for details";
        }

        return "CVE-only — select to load NVD description";
    }

    public static (IReadOnlyList<string> ReferenceUrls, string Description) ParseNvdEnrichment(string? enrichment)
    {
        if (string.IsNullOrWhiteSpace(enrichment))
            return ([], "");

        var urls = new List<string>();
        var textParts = new List<string>();
        foreach (var part in enrichment.Split(" | ", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (LooksLikeHttpUrl(part))
                urls.Add(part);
            else
                textParts.Add(part);
        }

        return (urls, string.Join(" | ", textParts));
    }

    private static bool LooksLikeHttpUrl(string value) =>
        Uri.TryCreate(value, UriKind.Absolute, out var uri) &&
        (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);

    private static string Truncate(string value, int maxLength)
    {
        if (value.Length <= maxLength)
            return value;

        return value[..maxLength].TrimEnd() + "…";
    }
}
