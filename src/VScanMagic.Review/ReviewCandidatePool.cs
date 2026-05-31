using VScanMagic.Core.Models;
using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class ReviewCandidatePool
{
    public static IReadOnlyList<ReviewFinding> GetExportFindings(ReviewSession session) =>
        ReviewSessionRanker.GetExportFindings(session);

    public static IReadOnlyList<ReviewFinding> GetCandidates(ReviewSession session, string? sourceFilter = null) =>
        session.Findings
            .Where(f => !f.IncludeInExport && !f.ExcludedFromExport && !f.ConnectSecureSuppressed)
            .Where(f => string.IsNullOrWhiteSpace(sourceFilter) ||
                        VulnerabilitySourceHelper.Normalize(f.Source)
                            .Equals(sourceFilter, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(f => f.RiskScore)
            .ThenBy(f => f.OriginalRank)
            .ToList();

    public static int CountCandidates(ReviewSession session, string? sourceFilter = null) =>
        GetCandidates(session, sourceFilter).Count;

    public static bool IsApplicationReserve(ReviewFinding finding) =>
        VulnerabilitySourceHelper.IsApplication(finding.Source) &&
        !finding.IncludeInExport &&
        !finding.ExcludedFromExport;

    public static string SeverityLabel(ReviewFinding finding)
    {
        if (finding.Critical > 0) return "Critical";
        if (finding.High > 0) return "High";
        if (finding.Medium > 0) return "Medium";
        if (finding.Low > 0) return "Low";
        return "—";
    }

    public static int SeverityRank(ReviewFinding finding) =>
        finding.Critical > 0 ? 4 :
        finding.High > 0 ? 3 :
        finding.Medium > 0 ? 2 :
        finding.Low > 0 ? 1 : 0;
}
