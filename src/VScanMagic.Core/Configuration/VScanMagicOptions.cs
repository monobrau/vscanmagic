namespace VScanMagic.Core.Configuration;

public sealed class VScanMagicOptions
{
    public string AppName { get; set; } = "VScanMagic";
    public string Author { get; set; } = "River Run";

    public SeverityWeights SeverityWeights { get; set; } = new();
    public CvssEquivalent CvssEquivalent { get; set; } = new();
    public List<string> FilteredProducts { get; set; } = [];
    public List<string> EolProductPatterns { get; set; } =
    [
        "OS-OUT-OF-SUPPORT", "OS-OUT-OF-ACTIVE-SUPPORT", "OS-OUT-OF-SECURITY-SUPPORT",
        "END OF LIFE", "End of Life", "end-of-life", "out of support", "Out of Support"
    ];
    public Dictionary<string, List<string>> WindowsConsolidation { get; set; } = new()
    {
        ["Windows Server 2012 (all versions)"] = ["Windows Server 2012", "Windows Server 2012 R2"],
        ["Windows 11 (all versions)"] = ["Windows 11", "Windows 1122H2", "Windows 1123H2", "Windows 1124H2"],
        ["Windows 10 (all versions)"] = ["Windows 10", "Windows 1022H2"]
    };
    public double SyntheticEpssForNoEpss { get; set; } = 0.1;
    public int MinNetworkVulnsInTopN { get; set; } = 0;
    /// <summary>Legacy auto-reserve for critical/high CVE-only findings (0 = technician picks in review).</summary>
    public int MinHighSeverityCveInTopN { get; set; } = 0;
    /// <summary>Scales CVE-only severity onto the same range as multi-vuln product buckets.</summary>
    public double CveOnlyRiskScale { get; set; } = 1.5;
}

public sealed class SeverityWeights
{
    public double Critical { get; set; } = 0.90;
    public double High { get; set; } = 0.80;
    public double Medium { get; set; } = 0.50;
    public double Low { get; set; } = 0.30;
}

public sealed class CvssEquivalent
{
    public double Critical { get; set; } = 9.0;
    public double High { get; set; } = 7.0;
    public double Medium { get; set; } = 5.0;
    public double Low { get; set; } = 3.0;
}
