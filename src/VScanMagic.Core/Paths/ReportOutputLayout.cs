namespace VScanMagic.Core.Paths;

public sealed class ReportOutputLayout
{
    public required string OutputDirectory { get; init; }
    public required string TextOutputDirectory { get; init; }
    public bool UsesStructuredPaths { get; init; }
    public bool UsesMiscSubfolder { get; init; }
    public string? ReportsPathPartial { get; init; }
}
