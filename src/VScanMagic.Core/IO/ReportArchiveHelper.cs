using System.IO.Compression;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.IO;

public static class ReportArchiveHelper
{
    private static readonly string[] SidecarSuffixes = [".extracting", ".extracted"];

    /// <summary>
    /// ConnectSecure sometimes returns a zip archive containing the real report file,
    /// saved with an .xlsx/.docx extension. Real Office files also use zip — detect wrappers
    /// by missing [Content_Types].xml at the archive root.
    /// </summary>
    public static string EnsureExtractedReportFile(string path)
    {
        if (!File.Exists(path) || !LooksLikeZip(path))
            return path;

        using var archive = ZipFile.OpenRead(path);
        if (IsRealOfficeOpenXml(archive))
            return path;

        var expectedExt = Path.GetExtension(path);
        var inner = FindInnerReportEntry(archive, expectedExt);
        if (inner is null)
            return path;

        var temp = VScanMagicPaths.GetTempFile("reports", inner.Name);
        inner.ExtractToFile(temp, overwrite: true);
        return temp;
    }

    /// <summary>
    /// If <paramref name="path"/> is a wrapper zip, replace it with the extracted report file.
    /// </summary>
    public static void NormalizeDownloadedReportFile(string path)
    {
        if (!File.Exists(path) || !LooksLikeZip(path))
        {
            CleanupSidecarFiles(path);
            return;
        }

        var temp = path + ".extracting";
        using (var archive = ZipFile.OpenRead(path))
        {
            if (IsRealOfficeOpenXml(archive))
            {
                CleanupSidecarFiles(path);
                return;
            }

            var inner = FindInnerReportEntry(archive, Path.GetExtension(path));
            if (inner is null)
            {
                CleanupSidecarFiles(path);
                return;
            }

            if (File.Exists(temp))
                File.Delete(temp);

            inner.ExtractToFile(temp, overwrite: true);
        }

        try
        {
            if (File.Exists(path))
                File.Delete(path);

            File.Move(temp, path);
        }
        finally
        {
            CleanupSidecarFiles(path);
        }
    }

    public static void CleanupSidecarFiles(string path)
    {
        foreach (var suffix in SidecarSuffixes)
        {
            var sidecar = path + suffix;
            if (File.Exists(sidecar))
            {
                try { File.Delete(sidecar); }
                catch { /* best effort */ }
            }
        }
    }

    private static bool IsRealOfficeOpenXml(ZipArchive archive) =>
        archive.Entries.Any(e =>
            e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));

    private static ZipArchiveEntry? FindInnerReportEntry(ZipArchive archive, string expectedExt) =>
        archive.Entries
            .Where(e => !string.IsNullOrEmpty(e.Name))
            .FirstOrDefault(e => e.Name.EndsWith(expectedExt, StringComparison.OrdinalIgnoreCase))
        ?? archive.Entries.FirstOrDefault(e => !string.IsNullOrEmpty(e.Name));

    private static bool LooksLikeZip(string path)
    {
        Span<byte> header = stackalloc byte[4];
        using var fs = File.OpenRead(path);
        if (fs.Read(header) < 4)
            return false;

        return header[0] == 0x50 && header[1] == 0x4B &&
               (header[2] == 0x03 || header[2] == 0x05 || header[2] == 0x07);
    }
}
