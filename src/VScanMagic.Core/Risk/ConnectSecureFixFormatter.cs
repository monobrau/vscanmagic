using System.Text.Json;
using System.Text.RegularExpressions;

namespace VScanMagic.Core.Risk;

public static partial class ConnectSecureFixFormatter
{
    private static readonly Regex PlaceholderPattern = PlaceholderFixRegex();
    private static readonly Regex QuotedValuePattern = QuotedFixValuesRegex();

    public static bool IsPlaceholder(string? rawFix)
    {
        if (string.IsNullOrWhiteSpace(rawFix))
            return true;

        var trimmed = rawFix.Trim();
        return PlaceholderPattern.IsMatch(trimmed)
               || trimmed.Equals("['None']", StringComparison.OrdinalIgnoreCase)
               || trimmed.Equals("[\"None\"]", StringComparison.OrdinalIgnoreCase);
    }

    public static string ToReadableText(string? rawFix)
    {
        if (string.IsNullOrWhiteSpace(rawFix) || IsPlaceholder(rawFix))
            return "";

        var items = ExtractFixItems(rawFix);
        if (items.Count == 0)
            return StripBracketNoise(rawFix);

        var parts = new List<string>();
        foreach (var item in items)
        {
            if (string.IsNullOrWhiteSpace(item) || item.Equals("None", StringComparison.OrdinalIgnoreCase))
                continue;

            if (KbNumberPattern().IsMatch(item))
                parts.Add($"Apply Windows Update KB{item}");
            else if (VersionNumberPattern().IsMatch(item))
                parts.Add($"Update to version {item} or later");
            else if (item.Equals("Latest Patch", StringComparison.OrdinalIgnoreCase))
                parts.Add("Apply the latest patch");
            else
                parts.Add(ProductConsolidator.GetProductMajorVersion(ProductNameNormalizer.Normalize(item)));
        }

        if (parts.Count == 0)
            return StripBracketNoise(rawFix);

        return string.Join(". ", parts.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static List<string> ExtractFixItems(string rawFix)
    {
        var trimmed = rawFix.Trim();
        if (trimmed.StartsWith('['))
        {
            var fromJson = TryParseJsonArray(trimmed);
            if (fromJson.Count > 0)
                return fromJson;
        }

        var fromQuotes = QuotedValuePattern.Matches(trimmed)
            .Select(m => m.Groups[1].Value.Trim())
            .Where(s => s.Length > 0)
            .ToList();

        return fromQuotes;
    }

    private static List<string> TryParseJsonArray(string value)
    {
        try
        {
            using var doc = JsonDocument.Parse(value);
            if (doc.RootElement.ValueKind != JsonValueKind.Array)
                return [];

            return doc.RootElement.EnumerateArray()
                .Where(e => e.ValueKind == JsonValueKind.String)
                .Select(e => e.GetString()?.Trim() ?? "")
                .Where(s => s.Length > 0)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static string StripBracketNoise(string rawFix)
    {
        var cleaned = rawFix.Trim().Trim('[', ']').Trim('"', '\'');
        return string.IsNullOrWhiteSpace(cleaned) ? "" : cleaned;
    }

    [GeneratedRegex(@"^\s*(\[?\s*)?(None|N/A|nil)\s*(\]?\s*)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
    private static partial Regex PlaceholderFixRegex();

    [GeneratedRegex("'([^']*)'", RegexOptions.Compiled)]
    private static partial Regex QuotedFixValuesRegex();

    [GeneratedRegex(@"^\d{6,8}$", RegexOptions.Compiled)]
    private static partial Regex KbNumberPattern();

    [GeneratedRegex(@"^[\d\.]+$", RegexOptions.Compiled)]
    private static partial Regex VersionNumberPattern();
}
