using Microsoft.Extensions.DependencyInjection;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;

namespace VScanMagic.Review;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddVScanMagicReview(this IServiceCollection services)
    {
        services.AddSingleton<IReviewSessionRepository, SqliteReviewSessionRepository>();
        services.AddSingleton<ReviewSessionFactory>();
        services.AddSingleton<CveEnrichmentService>();
        return services;
    }
}
