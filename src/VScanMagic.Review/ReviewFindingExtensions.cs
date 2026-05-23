using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class ReviewFindingExtensions
{
    public static IEnumerable<ReviewAffectedSystem> IncludedSystems(this ReviewFinding finding) =>
        (finding.AffectedSystems ?? []).Where(s => !s.ExcludedFromExport);
}
