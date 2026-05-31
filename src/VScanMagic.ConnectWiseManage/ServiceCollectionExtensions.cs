using Microsoft.Extensions.DependencyInjection;

namespace VScanMagic.ConnectWiseManage;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicConnectWiseManage(this IServiceCollection services)
    {
        services.AddHttpClient<ConnectWiseManageClient>();
        services.AddSingleton<ConnectWiseManageSettingsStore>();
        services.AddSingleton<ConnectWiseCompanyMapService>();
        services.AddSingleton<ConnectWiseManageTicketService>();
        return services;
    }
}
