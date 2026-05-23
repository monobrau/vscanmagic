using VScanMagic.Core.Configuration;
using VScanMagic.Core.Risk;
using VScanMagic.Review.Models;

namespace VScanMagic.Review;

public static class ReviewSessionRanker
{
    public static void EnsureLegacyFields(ReviewSession session, VScanMagicOptions? options = null)
    {
        session.Findings ??= [];
        var needsLegacy = session.Findings.Any(f => f.OriginalRank <= 0);
        if (needsLegacy)
        {
            foreach (var finding in session.Findings)
            {
                if (finding.OriginalRank <= 0)
                    finding.OriginalRank = finding.Rank > 0 ? finding.Rank : 1;
            }

            if (session.ExportTopN <= 0)
            {
                var included = session.Findings.Count(f => f.IncludeInExport);
                session.ExportTopN = included > 0 ? included : 10;
            }
        }

        foreach (var finding in session.Findings)
        {
            var cleaned = ProductNameNormalizer.FormatDisplayName(finding.Product, options);
            if (!string.IsNullOrWhiteSpace(cleaned))
                finding.Product = cleaned;

            if (!string.IsNullOrWhiteSpace(finding.OriginalFix))
            {
                var readableFix = ConnectSecureFixFormatter.ToReadableText(finding.OriginalFix);
                finding.OriginalFix = readableFix;
            }

            finding.CveIds = "";
        }

        session.Findings.RemoveAll(f => CveReferenceHelper.IsCveOnlyProduct(f.Product));

        ReviewFindingTimeEstimate.EnsureDefaults(session);
        Rebalance(session);
    }

    public static IReadOnlyList<ReviewFinding> GetExportFindings(ReviewSession session) =>
        session.Findings
            .Where(f => f.IncludeInExport)
            .OrderBy(f => f.OriginalRank)
            .ToList();

    public static int? GetExportRank(ReviewSession session, ReviewFinding finding)
    {
        if (!finding.IncludeInExport)
            return null;

        var rank = 1;
        foreach (var f in session.Findings.OrderBy(x => x.OriginalRank))
        {
            if (!f.IncludeInExport) continue;
            if (ReferenceEquals(f, finding) || f.OriginalRank == finding.OriginalRank)
                return rank;
            rank++;
        }

        return null;
    }

    public static void Rebalance(ReviewSession session)
    {
        var target = session.ExportTopN <= 0 ? int.MaxValue : session.ExportTopN;
        var included = session.Findings.Where(f => f.IncludeInExport).OrderBy(f => f.OriginalRank).ToList();

        while (included.Count < target)
        {
            var next = session.Findings
                .Where(f => !f.IncludeInExport && !f.ExcludedFromExport)
                .OrderBy(f => f.OriginalRank)
                .FirstOrDefault();
            if (next is null)
                break;

            next.IncludeInExport = true;
            included.Add(next);
        }

        var exportOrder = session.Findings.Where(f => f.IncludeInExport).OrderBy(f => f.OriginalRank).ToList();
        for (var i = 0; i < exportOrder.Count; i++)
            exportOrder[i].Rank = i + 1;
    }
}
