namespace VScanMagic.ConnectSecure;

public sealed record PatchApplicationLoadOptions
{
    public bool PatchableOnly { get; init; } = true;
    public string SeverityFilter { get; init; } = "all";
    public bool HideEndOfLife { get; init; }
}
