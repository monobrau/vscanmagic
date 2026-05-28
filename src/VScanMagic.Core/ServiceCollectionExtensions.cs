using Microsoft.Extensions.DependencyInjection;
using VScanMagic.Core.Configuration;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Nvd;
using VScanMagic.Core.Services;

namespace VScanMagic.Core;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicCore(this IServiceCollection services)
    {
        services.AddSingleton(_ => new VScanMagicOptions());
        services.AddSingleton<SettingsService>();
        services.AddSingleton<CompanyFolderMapService>();
        services.AddSingleton<ReportFolderHistoryService>();
        services.AddSingleton<PatchActivityHistoryService>();
        services.AddSingleton<RmitPlusSettingsService>();
        services.AddSingleton<ReportPathResolver>();
        services.AddSingleton<RemediationRuleService>();
        services.AddSingleton<TemplatesService>();
        services.AddSingleton<CoveredSoftwareService>();
        services.AddSingleton<NvdCveLookupService>();
        return services;
    }
}
