using VScanMagic.Core.Models;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class DeliverableLinksResolver
{
    public static DeliverableLinks Resolve(ReviewSession session, UserSettings userSettings) =>
        new()
        {
            TopNReportUrl = TrimOrEmpty(session.TopNReportUrl),
            ReportsFolderUrl = TrimOrEmpty(session.ReportsFolderUrl),
            SchedulingLinkUrl = FirstNonEmpty(session.SchedulingLinkUrl, userSettings.SchedulingLinkUrl)
        };

    /// <summary>Fills empty session scheduling link from Settings. OneDrive links are pasted per quarter on Deliverables.</summary>
    public static void ApplyDefaultsToSession(ReviewSession session, UserSettings userSettings)
    {
        if (string.IsNullOrWhiteSpace(session.SchedulingLinkUrl))
            session.SchedulingLinkUrl = userSettings.SchedulingLinkUrl.Trim();
    }

    private static string TrimOrEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "" : value.Trim();

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }

        return "";
    }
}
