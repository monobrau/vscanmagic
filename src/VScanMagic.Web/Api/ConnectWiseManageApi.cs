using VScanMagic.ConnectWiseManage;
using VScanMagic.Review.Storage;

namespace VScanMagic.Web.Api;

public static class ConnectWiseManageApi
{
    public static void MapConnectWiseManageApi(this WebApplication app)
    {
        var manage = app.MapGroup("/api/manage");

        manage.MapGet("/boards", async (ConnectWiseManageClient client, CancellationToken ct) =>
        {
            if (LocalApiGuard.RequireLoopbackBind() is { } denied)
                return denied;

            if (!client.IsConfigured)
                return Results.BadRequest(new { error = "ConnectWise Manage is not configured on Settings." });

            var boards = await client.GetBoardsAsync(ct).ConfigureAwait(false);
            return Results.Ok(boards);
        });

        manage.MapGet("/boards/{boardId:int}/statuses", async (
            int boardId,
            ConnectWiseManageClient client,
            CancellationToken ct) =>
        {
            if (LocalApiGuard.RequireLoopbackBind() is { } denied)
                return denied;

            if (boardId <= 0)
                return Results.BadRequest(new { error = "boardId is required." });

            if (!client.IsConfigured)
                return Results.BadRequest(new { error = "ConnectWise Manage is not configured on Settings." });

            var statuses = await client.GetBoardStatusesAsync(boardId, ct).ConfigureAwait(false);
            return Results.Ok(statuses);
        });

        var sessions = app.MapGroup("/api/sessions/{sessionId}");

        sessions.MapPost("/manage/tickets", async (
            string sessionId,
            IReviewSessionRepository sessionRepo,
            ConnectWiseManageTicketService ticketService,
            CancellationToken ct) =>
        {
            if (LocalApiGuard.RequireLoopbackBind() is { } denied)
                return denied;

            var session = await sessionRepo.GetAsync(sessionId, ct).ConfigureAwait(false);
            if (session is null)
                return Results.NotFound();

            try
            {
                var result = await ticketService.CreateTicketsForSessionAsync(session, ct).ConfigureAwait(false);
                return Results.Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return Results.BadRequest(new { error = ex.Message });
            }
        });

        sessions.MapPost("/manage/tickets/refresh", async (
            string sessionId,
            IReviewSessionRepository sessionRepo,
            ConnectWiseManageTicketService ticketService,
            CancellationToken ct) =>
        {
            if (LocalApiGuard.RequireLoopbackBind() is { } denied)
                return denied;

            var session = await sessionRepo.GetAsync(sessionId, ct).ConfigureAwait(false);
            if (session is null)
                return Results.NotFound();

            try
            {
                var count = await ticketService.RefreshTicketsForSessionAsync(session, ct).ConfigureAwait(false);
                return Results.Ok(new { refreshed = count });
            }
            catch (InvalidOperationException ex)
            {
                return Results.BadRequest(new { error = ex.Message });
            }
        });
    }
}
