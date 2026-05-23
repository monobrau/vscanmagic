using Microsoft.Extensions.DependencyInjection;
using VScanMagic.Core.Configuration;
using VScanMagic.Core.Models;
using VScanMagic.Data.Parsing;
using VScanMagic.Data.Scoring;

namespace VScanMagic.Data;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicData(this IServiceCollection services)
    {
        services.AddSingleton<ExcelVulnerabilityReader>();
        services.AddSingleton<TopVulnerabilityScorer>();
        services.AddSingleton<HostVulnerabilitySummarizer>();
        return services;
    }
}

public sealed class VulnerabilityPipeline(
    ExcelVulnerabilityReader reader,
    TopVulnerabilityScorer scorer,
    HostVulnerabilitySummarizer hostSummarizer)
{
    public PipelineResult ProcessFile(string excelPath, ReportFilters filters)
    {
        var records = reader.ReadFromFile(excelPath);
        var top = scorer.GetTopVulnerabilities(records, filters);
        return new PipelineResult(records, top);
    }

    public HostSummaryResult SummarizeHosts(string excelPath)
    {
        var records = reader.ReadFromFile(excelPath);
        return new HostSummaryResult(records, hostSummarizer.Summarize(records));
    }
}

public sealed record HostSummaryResult(
    IReadOnlyList<Core.Models.VulnerabilityRecord> AllRecords,
    IReadOnlyList<Core.Models.HostVulnerabilitySummary> Hosts);

public sealed record PipelineResult(
    IReadOnlyList<Core.Models.VulnerabilityRecord> AllRecords,
    IReadOnlyList<Core.Models.TopVulnerability> TopVulnerabilities);
