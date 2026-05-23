namespace VScanMagic.Core.Models;

public sealed class CoveredSoftwareEntry
{
    public string Pattern { get; set; } = "";
    public bool IsPattern { get; set; } = true;
    public bool Override { get; set; }
}
