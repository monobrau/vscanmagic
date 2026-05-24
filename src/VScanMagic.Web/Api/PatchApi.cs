using VScanMagic.ConnectSecure;

namespace VScanMagic.Web.Api;

public static class PatchApi
{
    public static void MapPatchApi(this WebApplication app)
    {
        app.MapGet("/api/connectsecure/patch/applications", async (
            int companyId,
            bool patchableOnly,
            ConnectSecurePatchService patchService,
            CancellationToken ct) =>
        {
            if (companyId <= 0)
                return Results.BadRequest(new { error = "companyId is required." });

            var apps = await patchService.GetPatchableApplicationsAsync(companyId, patchableOnly, ct);
            return Results.Ok(apps);
        });

        app.MapGet("/api/connectsecure/patch/assets", async (
            int companyId,
            int solutionId,
            ConnectSecurePatchService patchService,
            CancellationToken ct) =>
        {
            if (companyId <= 0)
                return Results.BadRequest(new { error = "companyId is required." });
            if (solutionId <= 0)
                return Results.BadRequest(new { error = "solutionId is required." });

            var assets = await patchService.GetPatchingAssetDetailsAsync(companyId, solutionId, ct);
            return Results.Ok(assets);
        });

        app.MapPost("/api/connectsecure/patch/now", async (
            ApplicationPatchRequest request,
            ConnectSecurePatchService patchService,
            CancellationToken ct) =>
        {
            try
            {
                var result = await patchService.PatchApplicationsNowAsync(request, ct);
                return Results.Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return Results.BadRequest(new { error = ex.Message });
            }
        });

        app.MapPost("/api/connectsecure/patch/schedule", async (
            ScheduledApplicationPatchRequest request,
            ConnectSecurePatchService patchService,
            CancellationToken ct) =>
        {
            try
            {
                var result = await patchService.ScheduleApplicationPatchAsync(request, ct);
                return Results.Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return Results.BadRequest(new { error = ex.Message });
            }
        });
    }
}
