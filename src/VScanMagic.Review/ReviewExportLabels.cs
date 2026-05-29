using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class ReviewExportLabels
{
    public static int GetExportCount(ReviewSession session) =>
        ReviewSessionRanker.GetExportFindings(session).Count;

    /// <summary>Short label for email, ticket notes, and section headings (e.g. "Top Ten", "Top 8").</summary>
    public static string GetTopNLabel(ReviewSession session)
    {
        var exportCount = GetExportCount(session);
        if (session.ExportTopN <= 0)
            return "Top";

        if (session.ExportTopN == 10 && exportCount == 10)
            return "Top Ten";

        return $"Top {exportCount}";
    }

    /// <summary>Full report title for docx cover and filename (e.g. "Top Ten Vulnerabilities Report").</summary>
    public static string GetReportTitle(ReviewSession session)
    {
        if (session.ExportTopN <= 0)
            return "Top Vulnerabilities Report";

        var label = GetTopNLabel(session);
        return label == "Top"
            ? "Top Vulnerabilities Report"
            : $"{label} Vulnerabilities Report";
    }
}
