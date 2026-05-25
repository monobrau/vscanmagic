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
            finding.Source = VulnerabilitySourceHelper.Normalize(finding.Source);

            var cleaned = ProductNameNormalizer.FormatDisplayName(finding.Product, options);
            if (!string.IsNullOrWhiteSpace(cleaned))
                finding.Product = cleaned;

            if (!string.IsNullOrWhiteSpace(finding.OriginalFix))
            {
                var readableFix = ConnectSecureFixFormatter.ToReadableText(finding.OriginalFix);
                finding.OriginalFix = readableFix;
            }

            finding.CveIds = CveReferenceHelper.NormalizeFindingCveIds(finding.CveIds, finding.Product);
        }

        ReviewFindingTimeEstimate.EnsureDefaults(session);
        RefreshExportRanks(session);
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

    public static void PromoteToExport(ReviewSession session, ReviewFinding finding)
    {
        if (finding.ConnectSecureSuppressed)
            return;

        finding.IncludeInExport = true;
        finding.ExcludedFromExport = false;
        finding.ManuallyPromoted = true;
        RefreshExportRanks(session);
    }

    public static void SetExportIncluded(ReviewSession session, ReviewFinding finding, bool include)
    {
        if (finding.ConnectSecureSuppressed)
            return;

        if (include)
        {
            finding.IncludeInExport = true;
            finding.ExcludedFromExport = false;
        }
        else
        {
            finding.IncludeInExport = false;
            finding.ExcludedFromExport = true;
            finding.ManuallyPromoted = false;
        }

        Rebalance(session);
    }

    public static void Rebalance(ReviewSession session)
    {
        var target = session.ExportTopN <= 0 ? int.MaxValue : session.ExportTopN;

        while (session.Findings.Count(f => f.IncludeInExport && !f.ManuallyPromoted) < target)
        {
            var next = session.Findings
                .Where(f => !f.IncludeInExport &&
                            !f.ExcludedFromExport &&
                            !f.ConnectSecureSuppressed &&
                            VulnerabilitySourceHelper.IsApplication(f.Source))
                .OrderBy(f => f.OriginalRank)
                .FirstOrDefault();
            if (next is null)
                break;

            next.IncludeInExport = true;
            next.ManuallyPromoted = false;
        }

        RefreshExportRanks(session);
    }

    public static void RefreshExportRanks(ReviewSession session)
    {
        var exportOrder = session.Findings.Where(f => f.IncludeInExport).OrderBy(f => f.OriginalRank).ToList();
        for (var i = 0; i < exportOrder.Count; i++)
            exportOrder[i].Rank = i + 1;
    }

    public static void MarkConnectSecureSuppressed(
        ReviewSession session,
        ReviewFinding finding,
        string reason,
        string? comments)
    {
        finding.ConnectSecureSuppressed = true;
        finding.SuppressionReason = reason.Trim();
        finding.SuppressionComments = string.IsNullOrWhiteSpace(comments) ? null : comments.Trim();
        finding.SuppressedAt = DateTimeOffset.UtcNow;
        finding.Status = FindingStatus.WontFix;
        finding.IncludeInExport = false;
        finding.ExcludedFromExport = true;
        finding.ManuallyPromoted = false;
        Rebalance(session);
    }

    public static void MarkConnectSecureUnsuppressed(ReviewSession session, ReviewFinding finding)
    {
        finding.ConnectSecureSuppressed = false;
        finding.SuppressionReason = null;
        finding.SuppressionComments = null;
        finding.SuppressedAt = null;
        finding.ConnectSecureSuppressRecordId = null;
        finding.ExcludedFromExport = false;
        if (finding.Status == FindingStatus.WontFix)
            finding.Status = FindingStatus.Open;
        Rebalance(session);
    }
}

