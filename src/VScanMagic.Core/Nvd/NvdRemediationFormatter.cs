using System.Text.Json;

namespace VScanMagic.Core.Nvd;

public static class NvdRemediationFormatter
{
    public static string BuildSummary(JsonElement cveItem)
    {
        if (cveItem.ValueKind != JsonValueKind.Object ||
            !cveItem.TryGetProperty("cve", out var cve) ||
            cve.ValueKind != JsonValueKind.Object)
            return "";

        var description = GetEnglishDescription(cve);
        var referenceUrls = CollectReferenceUrls(cve);

        var parts = new List<string>();
        foreach (var url in referenceUrls.Take(3))
            parts.Add(url);

        if (!string.IsNullOrWhiteSpace(description))
        {
            var trimmed = description.Trim();
            if (trimmed.Length > 400)
                trimmed = trimmed[..400].TrimEnd() + "…";

            if (parts.Count > 0)
                parts.Add(trimmed);
            else
                parts = [trimmed];
        }

        var output = string.Join(" | ", parts);
        if (output.Length > 500)
            output = output[..500].TrimEnd() + "…";

        return output;
    }

    private static string GetEnglishDescription(JsonElement cve)
    {
        if (!cve.TryGetProperty("descriptions", out var descriptions) ||
            descriptions.ValueKind != JsonValueKind.Array)
            return "";

        foreach (var item in descriptions.EnumerateArray())
        {
            if (item.TryGetProperty("lang", out var lang) &&
                string.Equals(lang.GetString(), "en", StringComparison.OrdinalIgnoreCase) &&
                item.TryGetProperty("value", out var value))
            {
                return value.GetString() ?? "";
            }
        }

        return "";
    }

    private static List<string> CollectReferenceUrls(JsonElement cve)
    {
        var preferred = new List<string>();
        var fallback = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!cve.TryGetProperty("references", out var references) ||
            references.ValueKind != JsonValueKind.Array)
            return preferred;

        foreach (var reference in references.EnumerateArray())
        {
            if (!reference.TryGetProperty("url", out var urlElement))
                continue;

            var url = urlElement.GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(url) || !seen.Add(url))
                continue;

            var source = reference.TryGetProperty("source", out var sourceElement)
                ? sourceElement.GetString() ?? ""
                : "";

            if (IsPreferredReference(url, source))
                preferred.Add(url);
            else
                fallback.Add(url);
        }

        if (preferred.Count > 0)
            return preferred;

        return fallback;
    }

    private static bool IsPreferredReference(string url, string source)
    {
        if (url.Contains("patch", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("advisory", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("security", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("bulletin", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("vendor", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("mitre", StringComparison.OrdinalIgnoreCase) ||
            url.Contains("nist.gov", StringComparison.OrdinalIgnoreCase))
            return true;

        return source.Contains("patch", StringComparison.OrdinalIgnoreCase) ||
               source.Contains("advisory", StringComparison.OrdinalIgnoreCase) ||
               source.Contains("security", StringComparison.OrdinalIgnoreCase) ||
               source.Contains("vendor", StringComparison.OrdinalIgnoreCase);
    }
}
