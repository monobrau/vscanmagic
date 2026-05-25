namespace VScanMagic.Core.Risk;

/// <summary>
/// Products that normally self-update and should use verify-first remediation guidance.
/// </summary>
public static class AutoUpdateSoftwareHelper
{
    private static readonly string[] DefaultPatterns =
    [
        "Google Chrome",
        "Mozilla Firefox",
        "Microsoft Edge",
        "Opera",
        "Brave",
        "Vivaldi"
    ];

    public static bool IsAutoUpdating(string? productName) =>
        IsAutoUpdating(productName, DefaultPatterns);

    public static bool IsAutoUpdating(string? productName, IEnumerable<string> patterns)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return false;

        return patterns.Any(pattern =>
            productName.Contains(pattern, StringComparison.OrdinalIgnoreCase));
    }
}
