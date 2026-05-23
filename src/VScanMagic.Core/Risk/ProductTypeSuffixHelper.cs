namespace VScanMagic.Core.Risk;

public static class ProductTypeSuffixHelper
{
    private static readonly string[] MicrosoftAppPatterns =
    [
        "Microsoft Office", "Microsoft 365", "Microsoft Teams", "Microsoft Edge",
        "Microsoft OneDrive", "Microsoft Outlook", "Microsoft Word", "Microsoft Excel",
        "Microsoft PowerPoint", "Microsoft Access", "Microsoft Publisher", "Microsoft Visio",
        "Microsoft Project", "Microsoft SharePoint", "Skype for Business",
        "Microsoft SQL Server Management Studio", "Microsoft Visual Studio Code",
        "Microsoft .NET Framework", "Microsoft .NET Core", "Microsoft .NET Runtime"
    ];

    private static readonly string[] VmwarePatterns =
    [
        "VMware", "VMWare", "vSphere", "vCenter", "ESXi", "VMware Tools",
        "VMware Workstation", "VMware Player", "VMware Horizon", "vRealize", "vCloud", "NSX"
    ];

    private static readonly string[] AutoUpdatePatterns = ["Google Chrome", "Mozilla Firefox"];

    public static string GetSuffix(string? productName, bool isRmitPlus = false)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return " - Update Required";

        if (Contains(productName, "Windows Server 2012") ||
            Contains(productName, "end-of-life") ||
            Contains(productName, "out of support"))
            return " - End of Support Migration Required";

        if (Contains(productName, "Windows 10"))
            return " - Windows 10 is End of Life";

        if (Contains(productName, "Windows 11"))
            return " - Updates Required";

        if (Contains(productName, "Windows Server"))
            return " - Updates Required";

        if (Contains(productName, "Windows"))
            return " - Patch Management Required";

        if (Contains(productName, "printer") || Contains(productName, "Ripple20"))
            return " - Firmware Update Required";

        if (Contains(productName, "Microsoft Teams"))
            return " - Application Update Required";

        if (isRmitPlus && IsMicrosoftApplication(productName))
            return " - Updates Required";

        if (isRmitPlus && IsVmwareProduct(productName))
            return " - Update Required";

        if (IsAutoUpdatingSoftware(productName))
            return " - This software updates automatically";

        return " - Update Required";
    }

    private static bool IsMicrosoftApplication(string productName) =>
        MicrosoftAppPatterns.Any(p => Contains(productName, p));

    private static bool IsVmwareProduct(string productName) =>
        VmwarePatterns.Any(p => Contains(productName, p));

    private static bool IsAutoUpdatingSoftware(string productName) =>
        AutoUpdatePatterns.Any(p => Contains(productName, p));

    private static bool Contains(string haystack, string needle) =>
        haystack.Contains(needle, StringComparison.OrdinalIgnoreCase);
}
