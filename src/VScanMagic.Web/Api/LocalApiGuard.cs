using VScanMagic.Web.Services;

namespace VScanMagic.Web.Api;

internal static class LocalApiGuard
{
    public static IResult? RequireLoopbackBind()
    {
        if (AppRestartService.IsRestartPermitted())
            return null;

        return Results.Json(
            new { error = "This endpoint is only available when VScanMagic is bound to loopback (127.0.0.1, localhost, or ::1)." },
            statusCode: StatusCodes.Status403Forbidden);
    }
}
