using Microsoft.Extensions.DependencyInjection;
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
        services.AddSingleton<VulnerabilitySupplementMerger>();
        services.AddSingleton<VulnerabilityPipeline>();
        return services;
    }
}
