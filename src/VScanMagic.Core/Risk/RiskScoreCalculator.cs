using VScanMagic.Core.Configuration;

namespace VScanMagic.Core.Risk;

public static class RiskScoreCalculator
{
    public static double GetAverageCvss(int critical, int high, int medium, int low, CvssEquivalent equiv)
    {
        var total = critical + high + medium + low;
        if (total == 0) return 0;
        var weighted = critical * equiv.Critical + high * equiv.High + medium * equiv.Medium + low * equiv.Low;
        return Math.Round(weighted / total, 2);
    }

    public static double GetCompositeRiskScore(
        int critical, int high, int medium, int low,
        double epssScore, string productName, int vulnCount,
        VScanMagicOptions options)
    {
        var severityWeightedSum =
            critical * options.SeverityWeights.Critical +
            high * options.SeverityWeights.High +
            medium * options.SeverityWeights.Medium +
            low * options.SeverityWeights.Low;

        if (IsEolProduct(productName, options))
            severityWeightedSum += vulnCount * 1.0;

        var effectiveEpss = epssScore;
        if (double.IsNaN(effectiveEpss))
            effectiveEpss = options.SyntheticEpssForNoEpss;

        var epssFactor = 1.0 + effectiveEpss;
        return Math.Round(severityWeightedSum * epssFactor, 2);
    }

    public static bool IsEolProduct(string productName, VScanMagicOptions options)
    {
        if (string.IsNullOrWhiteSpace(productName)) return false;
        return options.EolProductPatterns.Any(p =>
            productName.Contains(p, StringComparison.OrdinalIgnoreCase));
    }

    public static (string BackgroundHex, string TextHex, string Name) GetRiskColor(double riskScore, double maxRiskScore)
    {
        if (maxRiskScore <= 0) maxRiskScore = 10;
        var thresholds = new[]
        {
            (Pct: 1.00, Bg: "DC143C", Fg: "FFFFFF", Name: "Critical"),
            (Pct: 0.70, Bg: "FF4500", Fg: "FFFFFF", Name: "Very High"),
            (Pct: 0.50, Bg: "FF8C00", Fg: "FFFFFF", Name: "High"),
            (Pct: 0.30, Bg: "FFA500", Fg: "000000", Name: "Medium-High"),
            (Pct: 0.00, Bg: "FFFF00", Fg: "000000", Name: "Medium")
        };

        foreach (var t in thresholds)
        {
            if (riskScore >= maxRiskScore * t.Pct)
                return (t.Bg, t.Fg, t.Name);
        }

        return ("FFFF00", "000000", "Medium");
    }
}
