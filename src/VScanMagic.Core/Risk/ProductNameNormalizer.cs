using System.Text.Json;
using System.Text.RegularExpressions;
using VScanMagic.Core.Configuration;

namespace VScanMagic.Core.Risk;

public static partial class ProductNameNormalizer
{
    private static readonly Regex QuotedStringPattern = QuotedStringsRegex();

    public static string Normalize(string? product)
    {
        var names = ParseProductNames(product);
        if (names.Count == 0)
            return "";

        if (names.Count == 1)
            return names[0];

        return string.Join(", ", names);
    }

    public static IReadOnlyList<string> ParseProductNames(string? product)
    {
        if (string.IsNullOrWhiteSpace(product))
            return [];

        var trimmed = product.Trim();
        if (trimmed.StartsWith('['))
        {
            var fromJson = TryParseJsonArray(trimmed);
            if (fromJson.Count > 0)
                return fromJson;

            var fromQuotes = ExtractQuotedStrings(trimmed);
            if (fromQuotes.Count > 0)
                return fromQuotes;
        }

        return [CleanLiteral(trimmed)];
    }

    public static string FormatDisplayName(string? product, VScanMagicOptions? options = null)
    {
        var normalized = Normalize(product);
        if (string.IsNullOrWhiteSpace(normalized))
            return "";

        var parts = normalized.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length <= 1)
        {
            var single = ProductConsolidator.GetProductMajorVersion(parts.Length == 1 ? parts[0] : normalized);
            return options is null ? single : ProductConsolidator.GetConsolidatedProduct(single, options);
        }

        var cleaned = parts
            .Select(p =>
            {
                var major = ProductConsolidator.GetProductMajorVersion(p);
                return options is null ? major : ProductConsolidator.GetConsolidatedProduct(major, options);
            })
            .Distinct(StringComparer.OrdinalIgnoreCase);

        return string.Join(", ", cleaned);
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
                .Select(e => CleanLiteral(e.GetString() ?? ""))
                .Where(s => s.Length > 0)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static List<string> ExtractQuotedStrings(string value)
    {
        return QuotedStringPattern.Matches(value)
            .Select(m => CleanLiteral(m.Groups[1].Value.Replace("\\\"", "\"")))
            .Where(s => s.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string CleanLiteral(string value)
    {
        var s = value.Trim().Trim('"');
        if (s.StartsWith('[') && s.EndsWith(']'))
            s = Normalize(s);
        return s.Trim();
    }

    [GeneratedRegex("\"((?:\\\\.|[^\"\\\\])*)\"", RegexOptions.Compiled)]
    private static partial Regex QuotedStringsRegex();
}
