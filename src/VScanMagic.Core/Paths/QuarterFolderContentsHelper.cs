namespace VScanMagic.Core.Paths;

/// <summary>
/// Manages quarter folder contents for replace-before-rerun workflows (OneDrive link stability).
/// </summary>
public static class QuarterFolderContentsHelper
{
    public static bool HasContent(string quarterDirectory)
    {
        if (string.IsNullOrWhiteSpace(quarterDirectory) || !Directory.Exists(quarterDirectory))
            return false;

        return CountFiles(quarterDirectory) > 0;
    }

    public static int CountFiles(string quarterDirectory)
    {
        if (string.IsNullOrWhiteSpace(quarterDirectory) || !Directory.Exists(quarterDirectory))
            return 0;

        var full = Path.GetFullPath(quarterDirectory.Trim());
        var count = Directory.GetFiles(full, "*", SearchOption.TopDirectoryOnly).Length;
        var misc = Path.Combine(full, ReportPathResolver.MiscSubfolderName);
        if (Directory.Exists(misc))
            count += Directory.GetFiles(misc, "*", SearchOption.TopDirectoryOnly).Length;

        return count;
    }

    /// <summary>Removes files from the quarter folder and Misc subfolder; keeps folder structure.</summary>
    public static void ClearContents(string quarterDirectory)
    {
        if (string.IsNullOrWhiteSpace(quarterDirectory) || !Directory.Exists(quarterDirectory))
            return;

        var full = Path.GetFullPath(quarterDirectory.Trim());
        DeleteFilesInDirectory(full);

        var misc = Path.Combine(full, ReportPathResolver.MiscSubfolderName);
        if (Directory.Exists(misc))
            DeleteFilesInDirectory(misc);
    }

    private static void DeleteFilesInDirectory(string directory)
    {
        foreach (var file in Directory.GetFiles(directory))
        {
            try
            {
                File.Delete(file);
            }
            catch
            {
                // Best-effort clear before download; locked files surface on write.
            }
        }
    }
}
