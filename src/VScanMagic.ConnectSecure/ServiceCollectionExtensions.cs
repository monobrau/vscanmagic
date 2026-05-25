using Microsoft.Extensions.DependencyInjection;

namespace VScanMagic.ConnectSecure;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicConnectSecure(this IServiceCollection services)
    {
        services.AddSingleton<ConnectSecureOptions>();
        services.AddSingleton<RateLimiter>();
        services.AddSingleton<ConnectSecureCacheService>();
        // Singleton so Configure() in Settings/Home applies to ConnectSecureReportService too
        // (typed HttpClient registration creates a separate instance per injection site).
        services.AddSingleton<ConnectSecureClient>(sp =>
        {
            var http = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(90)
            };
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "VScanMagic/1.0");
            return new ConnectSecureClient(
                http,
                sp.GetRequiredService<RateLimiter>(),
                sp.GetRequiredService<ConnectSecureOptions>(),
                sp.GetRequiredService<ConnectSecureCacheService>());
        });
        services.AddSingleton<ConnectSecureReportService>();
        services.AddSingleton<ConnectSecureCompanyReviewService>();
        services.AddSingleton<ConnectSecureAgentService>();
        services.AddSingleton<ConnectSecureDiscoverySettingsService>();
        services.AddSingleton<ConnectSecureCompanyCredentialsService>();
        services.AddSingleton<ConnectSecureIntegrationService>();
        services.AddSingleton<ConnectSecureProbeConfigurationService>();
        services.AddSingleton<ConnectSecurePatchService>();
        services.AddSingleton<ConnectSecureSuppressService>();
        services.AddSingleton<ConnectSecureReviewSuppressService>();
        services.AddSingleton<ConnectSecureScanService>();
        return services;
    }
}
