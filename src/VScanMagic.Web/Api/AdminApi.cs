using VScanMagic.Web.Services;

namespace VScanMagic.Web.Api;

public static class AdminApi
{
    public static void MapAdminApi(this WebApplication app)
    {
        app.MapPost("/api/admin/restart", async (AppRestartService restart, CancellationToken ct) =>
        {
            if (!AppRestartService.IsRestartPermitted())
                return Results.Json(new { error = "Restart is only allowed on loopback bind addresses." }, statusCode: StatusCodes.Status403Forbidden);

            await restart.ScheduleRestartAsync(ct).ConfigureAwait(false);
            return Results.Ok(new { message = "Restart scheduled. Reload the page in a few seconds." });
        });
    }
}
