namespace VScanMagic.ConnectSecure;

public enum ReportCatalogScope
{
    Company,
    Global
}

public sealed record CatalogReportDownloadRequest(string ReportId, string Name, string Extension, string? ReportType = null);

/// <summary>One report category from ConnectSecure with available export formats.</summary>
public sealed class ReportCatalogGroup
{
    public required string Name { get; init; }
    public IReadOnlyDictionary<string, CatalogFormatOption> Formats { get; init; } =
        new Dictionary<string, CatalogFormatOption>(StringComparer.OrdinalIgnoreCase);
}

public sealed class CatalogFormatOption
{
    public required string ReportId { get; init; }
    public required string Extension { get; init; }
}

public static class ReportCatalogBuilder
{
    private static readonly string[] FormatOrder = ["xlsx", "docx", "pdf"];

    public static IReadOnlyList<ReportCatalogGroup> BuildGroups(IReadOnlyList<StandardReportDescriptor> descriptors)
    {
        var byCategory = new Dictionary<string, Dictionary<string, CatalogFormatOption>>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in descriptors)
        {
            var ext = item.ReportType.Trim().ToLowerInvariant();
            if (ext is not ("xlsx" or "docx" or "pdf"))
                continue;

            var name = string.IsNullOrWhiteSpace(item.CategoryDisplay)
                ? item.DisplayName
                : item.CategoryDisplay;
            if (string.IsNullOrWhiteSpace(name))
                name = item.Category;
            if (string.IsNullOrWhiteSpace(name))
                name = "Report";

            if (!byCategory.TryGetValue(name, out var formats))
            {
                formats = new Dictionary<string, CatalogFormatOption>(StringComparer.OrdinalIgnoreCase);
                byCategory[name] = formats;
            }

            formats[ext] = new CatalogFormatOption { ReportId = item.Id, Extension = ext };
        }

        return byCategory
            .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
            .Select(kv => new ReportCatalogGroup { Name = kv.Key, Formats = kv.Value })
            .ToList();
    }

    public static IReadOnlyList<string> FormatOrderList => FormatOrder;
}
