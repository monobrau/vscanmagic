using VScanMagic.Core.Models;

namespace VScanMagic.Core.Paths;

public static class ReportOutputDirectoryPersistence
{
    /// <summary>
    /// When Reports Base Path is configured, output folders are resolved from that root each time.
    /// Do not persist deep quarter paths into LastOutputDirectory (they caused path nesting on the next run).
    /// </summary>
    public static void UpdateLastOutputDirectory(UserSettings settings, string outputDirectory)
    {
        if (!string.IsNullOrWhiteSpace(ReportPathResolver.NormalizeConfiguredBasePath(settings.ReportsBasePath)))
            return;

        settings.LastOutputDirectory = outputDirectory;
    }
}
