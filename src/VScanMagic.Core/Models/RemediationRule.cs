namespace VScanMagic.Core.Models;

public sealed class RemediationRule
{
    public string Pattern { get; set; } = "*";
    public string WordText { get; set; } = "";
    public string TicketText { get; set; } = "";
    public bool IsDefault { get; set; }
    public RemediationGuidanceStyle GuidanceStyle { get; set; } = RemediationGuidanceStyle.Standard;
}
