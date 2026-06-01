using Microsoft.Extensions.DependencyInjection;

namespace VScanMagic.ConnectWiseManage;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicConnectWiseManage(this IServiceCollection services)
    {
        // Singleton so Configure() in Settings applies to ConnectWiseManageTicketService too
        // (typed HttpClient registration creates a separate instance per injection site).
        services.AddSingleton<ConnectWiseManageClient>(sp =>
        {
            var http = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "VScanMagic/1.0");
            return new ConnectWiseManageClient(http);
        });
        services.AddSingleton<ConnectWiseManageSettingsStore>();
        services.AddSingleton<ConnectWiseCompanyMapService>();
        services.AddSingleton<ConnectWiseManageTicketService>();
        return services;
    }
}
