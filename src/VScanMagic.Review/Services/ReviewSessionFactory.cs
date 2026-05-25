using VScanMagic.Core.Configuration;
using VScanMagic.Core.Models;
using VScanMagic.Core.Risk;
using VScanMagic.Core.Services;
using VScanMagic.Review.Models;

namespace VScanMagic.Review.Services;

public sealed class ReviewSessionFactory(RemediationRuleService remediationRules, VScanMagicOptions options)
{
    public ReviewSession CreateFromScoredResult(
        string clientName,
        string scanDate,
        ScoredVulnerabilityResult scored,
        string presenter,
        string? sourceFilePath = null,
        string? companyId = null,
        int exportTopN = 10,
        bool isRmitPlus = false)
    {
        var autoExportKeys = scored.AutoExportApplication
            .Select(v => VulnerabilitySourceHelper.ExportKey(v.Source, v.Product))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var ordered = OrderForSession(scored.AllFiltered);

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
        foreach (var item in ordered)
        {
            var key = VulnerabilitySourceHelper.ExportKey(item.Source, item.Product);
            var includeInExport = autoExportKeys.Contains(key);

            session.Findings.Add(CreateFinding(item, originalRank, includeInExport, isRmitPlus));
            originalRank++;
        }

        ReviewSessionRanker.Rebalance(session);
        return session;
    }

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
        var autoKeys = topVulns
            .Where(v => VulnerabilitySourceHelper.IsApplication(v.Source))
            .Take(exportTopN <= 0 ? int.MaxValue : exportTopN)
            .Select(v => VulnerabilitySourceHelper.ExportKey(v.Source, v.Product))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var scored = new ScoredVulnerabilityResult
        {
            AllFiltered = topVulns,
            AutoExportApplication = topVulns
                .Where(v => autoKeys.Contains(VulnerabilitySourceHelper.ExportKey(v.Source, v.Product)))
                .ToList()
        };

        return CreateFromScoredResult(
            clientName, scanDate, scored, presenter, sourceFilePath, companyId, exportTopN, isRmitPlus);
    }

    private static IReadOnlyList<TopVulnerability> OrderForSession(IReadOnlyList<TopVulnerability> allFiltered)
    {
        var application = allFiltered
            .Where(v => VulnerabilitySourceHelper.IsApplication(v.Source))
            .OrderByDescending(v => v.RiskScore)
            .ToList();
        var network = allFiltered
            .Where(v => VulnerabilitySourceHelper.IsNetwork(v.Source))
            .OrderByDescending(v => v.RiskScore)
            .ToList();
        var registry = allFiltered
            .Where(v => VulnerabilitySourceHelper.IsRegistry(v.Source))
            .OrderByDescending(v => v.RiskScore)
            .ToList();
        var other = allFiltered
            .Where(v => !VulnerabilitySourceHelper.IsApplication(v.Source) &&
                        !VulnerabilitySourceHelper.IsNetwork(v.Source) &&
                        !VulnerabilitySourceHelper.IsRegistry(v.Source))
            .OrderByDescending(v => v.RiskScore)
            .ToList();

        var ordered = new List<TopVulnerability>(allFiltered.Count);
        ordered.AddRange(application);
        ordered.AddRange(network);
        ordered.AddRange(registry);
        ordered.AddRange(other);
        return ordered;
    }

    private ReviewFinding CreateFinding(
        TopVulnerability item,
        int originalRank,
        bool includeInExport,
        bool isRmitPlus)
    {
        var displayProduct = ProductNameNormalizer.FormatDisplayName(item.Product, options);
        var remediation = ResolveInitialRemediation(displayProduct, item.Fix);
        var technicalFix = ConnectSecureFixFormatter.IsPlaceholder(item.Fix)
            ? ""
            : ConnectSecureFixFormatter.ToReadableText(item.Fix);
        var groupKey = ProductConsolidator.GetTimeEstimateGroupKey(displayProduct);

        return new ReviewFinding
        {
            Rank = originalRank,
            OriginalRank = originalRank,
            Product = displayProduct,
            Source = VulnerabilitySourceHelper.Normalize(item.Source),
            RiskScore = item.RiskScore,
            Epss = item.EpssScore,
            AvgCvss = item.AvgCvss,
            VulnCount = item.VulnCount,
            Critical = item.Critical,
            High = item.High,
            Medium = item.Medium,
            Low = item.Low,
            CveIds = CveReferenceHelper.NormalizeFindingCveIds(item.CveIds, displayProduct),
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
        };
    }

    private string ResolveInitialRemediation(string displayProduct, string? rawFix)
    {
        if (!string.IsNullOrWhiteSpace(rawFix) && !ConnectSecureFixFormatter.IsPlaceholder(rawFix))
        {
            var readable = ConnectSecureFixFormatter.ToReadableText(rawFix);
            if (!string.IsNullOrWhiteSpace(readable))
                return readable;
        }

        return remediationRules.GetGuidance(displayProduct, forWord: true);
    }
}
