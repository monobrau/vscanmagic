using VScanMagic.Core.Configuration;
using VScanMagic.Core.Models;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Review.Services;

public sealed class ReviewSessionFactory(RemediationRuleService remediationRules, VScanMagicOptions options)
{
    public ReviewSession CreateFromTopVulnerabilities(
        string clientName,
        string scanDate,
        IReadOnlyList<TopVulnerability> topVulns,
        string presenter,
        string? sourceFilePath = null,
        string? companyId = null,
        int exportTopN = 10,
        bool isRmitPlus = false)
    {
        var session = new ReviewSession
        {
            ClientName = clientName,
            ScanDate = scanDate,
            Presenter = presenter,
            SourceFilePath = sourceFilePath,
            CompanyId = companyId,
            ExportTopN = exportTopN,
            IsRmitPlus = isRmitPlus
        };

        var originalRank = 1;
        foreach (var item in topVulns)
        {
            var displayProduct = ProductNameNormalizer.FormatDisplayName(item.Product, options);
            var remediation = ResolveInitialRemediation(displayProduct, item.Fix);
            var technicalFix = ConnectSecureFixFormatter.IsPlaceholder(item.Fix)
                ? ""
                : ConnectSecureFixFormatter.ToReadableText(item.Fix);
            var includeInExport = exportTopN <= 0 || originalRank <= exportTopN;
            var groupKey = ProductConsolidator.GetTimeEstimateGroupKey(displayProduct);

            session.Findings.Add(new ReviewFinding
            {
                Rank = includeInExport ? originalRank : originalRank,
                OriginalRank = originalRank,
                Product = displayProduct,
                Source = item.Source,
                RiskScore = item.RiskScore,
                Epss = item.EpssScore,
                AvgCvss = item.AvgCvss,
                VulnCount = item.VulnCount,
                Critical = item.Critical,
                High = item.High,
                Medium = item.Medium,
                Low = item.Low,
                CveIds = "",
                OriginalFix = technicalFix,
                OriginalRemediation = remediation,
                RevisedRemediation = remediation,
                IncludeInExport = includeInExport,
                ThirdParty = FirstPartyVendorHelper.IsThirdPartyByDefault(groupKey, isRmitPlus),
                TimeEstimateInitialized = true,
                AffectedSystems = (item.AffectedSystems ?? []).Select(s => new ReviewAffectedSystem
                {
                    HostName = s.HostName,
                    Ip = s.Ip,
                    Username = s.Username,
                    VulnCount = s.VulnCount
                }).ToList()
            });

            originalRank++;
        }

        ReviewSessionRanker.Rebalance(session);
        return session;
    }

    private string ResolveInitialRemediation(string displayProduct, string? rawFix)
    {
        if (remediationRules.TryGetSpecificGuidance(displayProduct, forWord: true, out var ruleText))
            return ruleText;

        if (!ConnectSecureFixFormatter.IsPlaceholder(rawFix))
        {
            var fromFix = ConnectSecureFixFormatter.ToReadableText(rawFix);
            if (!string.IsNullOrWhiteSpace(fromFix))
                return fromFix;
        }

        return remediationRules.GetGuidance(displayProduct, forWord: true);
    }
}
