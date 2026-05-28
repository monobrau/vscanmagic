namespace VScanMagic.Core.Models;

public sealed class ReportFilters
{
    public double MinEpss { get; set; }
    public bool IncludeCritical { get; set; } = true;
    public bool IncludeHigh { get; set; } = true;
    public bool IncludeMedium { get; set; }
    public bool IncludeLow { get; set; }
    public int TopN { get; set; } = 10;

    public static int ParseTopN(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Equals("All", StringComparison.OrdinalIgnoreCase))
            return 0;

        return int.TryParse(value, out var n) && n > 0 ? n : 10;
    }

    public static int CandidatePoolSize(int exportTopN) =>
        exportTopN <= 0 ? 0 : exportTopN + Math.Max(exportTopN, 10);

    public ReportFilters WithTopN(int topN) => new()
    {
        MinEpss = MinEpss,
        IncludeCritical = IncludeCritical,
        IncludeHigh = IncludeHigh,
        IncludeMedium = IncludeMedium,
        IncludeLow = IncludeLow,
        TopN = topN
    };

    public static ReportFilters FromUserSettings(UserSettings settings)
    {
        var exportTopN = ParseTopN(settings.FilterTopN);
        return new ReportFilters
        {
            MinEpss = settings.FilterMinEpss,
            IncludeCritical = settings.FilterIncludeCritical,
            IncludeHigh = settings.FilterIncludeHigh,
            IncludeMedium = settings.FilterIncludeMedium,
            IncludeLow = settings.FilterIncludeLow,
            TopN = exportTopN
        };
    }

    public static ReportFilters CandidatePoolFilters(UserSettings settings)
    {
        var export = FromUserSettings(settings);
        var poolSize = CandidatePoolSize(export.TopN);
        return poolSize <= 0 ? export : export.WithTopN(poolSize);
    }
}
