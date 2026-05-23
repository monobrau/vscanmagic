namespace VScanMagic.Core.Risk;

public static class FirstPartyVendorHelper
{
    private static readonly string[] FirstPartySubstrings =
    [
        "Sonicwall", "Fortinet", "FortiGate", "Forti",
        "Microsoft",
        "Windows 11", "Windows 10", "Windows Server", "Windows 8", "Windows 7",
        "HP Pro", "HP LaserJet", "HP OfficeJet", "Hewlett-Packard",
        "Duo Security", "Duo ",
        "VMware", "vSphere", "VMware Tools"
    ];

    public static bool IsFirstPartyVendor(string? productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return false;

        var p = productName.Trim();
        if (p.StartsWith("HP ", StringComparison.OrdinalIgnoreCase) ||
            p.Contains(" HP ", StringComparison.OrdinalIgnoreCase))
            return true;

        return FirstPartySubstrings.Any(s => p.Contains(s, StringComparison.OrdinalIgnoreCase));
    }

    public static bool IsThirdPartyByDefault(string? productName, bool isRmitPlus) =>
        isRmitPlus && !IsFirstPartyVendor(productName);

}
