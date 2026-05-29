using VScanMagic.Core.Models;
using VScanMagic.Core.Services;

namespace VScanMagic.Core.Paths;

/// <summary>
/// Keeps download, review, and export on the same quarter folder for a session.
/// </summary>
public static class SessionOutputLayoutResolver
{
    public static ReportOutputLayout ResolveForSession(
        ReportPathResolver pathResolver,
        UserSettings settings,
        string clientName,
        string scanDate,
        int companyId,
        string? sessionOutputDirectory,
        string? sourceFilePath,
        ReportFolderHistoryService? folderHistory = null,
        string? fallbackPath = null)
    {
        var defaultLayout = pathResolver.Resolve(
            settings,
            companyId,
            clientName,
            scanDate,
            fallbackPath);

        var pinnedPath = GetPinnedOutputDirectory(
            sessionOutputDirectory,
            sourceFilePath,
            clientName,
            folderHistory);

        if (string.IsNullOrWhiteSpace(pinnedPath) || !Directory.Exists(pinnedPath))
            return defaultLayout;

        var pinnedNorm = Path.GetFullPath(pinnedPath.Trim());
        if (PathsMatch(defaultLayout.OutputDirectory, pinnedNorm))
            return defaultLayout;

        return ReportPathResolver.LayoutForExistingDirectory(pinnedNorm, clientName);
    }

    public static string? GetPinnedOutputDirectory(
        string? sessionOutputDirectory,
        string? sourceFilePath,
        string clientName,
        ReportFolderHistoryService? folderHistory = null)
    {
        if (!string.IsNullOrWhiteSpace(sessionOutputDirectory))
        {
            var path = Path.GetFullPath(sessionOutputDirectory.Trim());
            if (Directory.Exists(path))
                return path;
        }

        var fromSource = InferQuarterDirectoryFromSourceFile(sourceFilePath);
        if (!string.IsNullOrWhiteSpace(fromSource) && Directory.Exists(fromSource))
            return fromSource;

        return folderHistory?.GetLatestOutputPath(clientName);
    }

    /// <summary>
    /// Quarter folder containing ConnectSecure downloads (parent when the source file lives in Misc).
    /// </summary>
    public static string? InferQuarterDirectoryFromSourceFile(string? sourceFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            return null;

        var dir = Path.GetDirectoryName(Path.GetFullPath(sourceFilePath));
        if (string.IsNullOrWhiteSpace(dir))
            return null;

        if (Path.GetFileName(dir).Equals(ReportPathResolver.MiscSubfolderName, StringComparison.OrdinalIgnoreCase))
            return Path.GetDirectoryName(dir);

        return dir;
    }

    private static bool PathsMatch(string left, string right) =>
        string.Equals(
            Path.GetFullPath(left.Trim()),
            Path.GetFullPath(right.Trim()),
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
}
