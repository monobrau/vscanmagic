namespace VScanMagic.Core.Paths;

public static class VScanMagicPaths
{
    public static string GetConfigDirectory(string? overrideDirectory = null)
    {
        if (!string.IsNullOrWhiteSpace(overrideDirectory))
            return overrideDirectory;

        if (OperatingSystem.IsWindows())
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VScanMagic");

        var xdg = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME");
        if (!string.IsNullOrWhiteSpace(xdg))
            return Path.Combine(xdg, "vscanmagic");

        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config", "vscanmagic");
    }

    public static string SettingsFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_Settings.json");

    public static string RemediationRulesFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_RemediationRules.json");

    public static string ConnectSecureCredentialsFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "ConnectSecure-Credentials.json");

    public static string ConnectWiseManageCredentialsFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "ConnectWise-Manage-Credentials.json");

    public static string ConnectWiseManageOptionsFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "ConnectWise-Manage-Options.json");

    public static string ConnectWiseCompanyMapFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "ConnectWise_CompanyMap.json");

    public static string TemplatesFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_Templates.json");

    public static string CompanyFolderMapFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_CompanyFolderMap.json");

    public static string ReportFolderHistoryFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_ReportFolderHistory.json");

    public static string PatchActivityHistoryFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_PatchActivityHistory.json");

    public static string RmitPlusSettingsFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_RMITPlusSettings.json");

    public static string CoveredSoftwareFile(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "VScanMagic_CoveredSoftware.json");

    public static string SessionsDirectory(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "sessions");

    public static string TemplatesDirectory(string? configDir = null) =>
        Path.Combine(GetConfigDirectory(configDir), "templates");

    public static string GetTempDirectory()
    {
        var root = Path.Combine(Path.GetTempPath(), "VScanMagic");
        Directory.CreateDirectory(root);
        return root;
    }

    public static string GetTempFile(string subfolder, string baseName)
    {
        var dir = Path.Combine(GetTempDirectory(), subfolder);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, $"{Guid.NewGuid():N}_{baseName}");
    }

    public static string NvdCacheDirectory()
    {
        var dir = Path.Combine(GetConfigDirectory(), "nvd-cache");
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string NvdCacheFile(string cveId) =>
        Path.Combine(NvdCacheDirectory(), $"{cveId.Trim().ToUpperInvariant()}.json");
}
