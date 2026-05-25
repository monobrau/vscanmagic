using System.Text.RegularExpressions;
using VScanMagic.Core.Models;
using VScanMagic.Core.Services;

namespace VScanMagic.Core.Paths;

public sealed class ReportPathResolver(CompanyFolderMapService companyFolderMapService)
{
    private const int MaxPathLength = 250;
    private const string VulnerabilityScansSegment = "Network Documentation/Vulnerability Scans";

    public ReportOutputLayout Resolve(UserSettings settings, int companyId, string companyName, string scanDate, string? fallbackPath = null)
    {
        var usesMisc = !string.IsNullOrWhiteSpace(settings.ReportsBasePath);
        var fallback = ResolveFallbackDirectory(settings, fallbackPath);
        var displayName = string.IsNullOrWhiteSpace(companyName) ? "Client" : companyName.Trim();

        var basePath = settings.ReportsBasePath?.Trim();
        if (string.IsNullOrWhiteSpace(basePath) || !Directory.Exists(basePath))
        {
            if (usesMisc)
                return BuildClientQuarterLayout(fallback, displayName, scanDate);

            EnsureDirectory(fallback);
            return new ReportOutputLayout
            {
                OutputDirectory = fallback,
                TextOutputDirectory = fallback,
                UsesStructuredPaths = false,
                UsesMiscSubfolder = false,
                ReportsPathPartial = null
            };
        }

        basePath = Path.GetFullPath(basePath);

        if (TryResolveStructuredPath(basePath, companyId, displayName, scanDate, out var structuredPath))
        {
            EnsureDirectory(structuredPath);

            return new ReportOutputLayout
            {
                OutputDirectory = structuredPath,
                TextOutputDirectory = structuredPath,
                UsesStructuredPaths = true,
                UsesMiscSubfolder = false,
                ReportsPathPartial = GetReportsPathPartial(structuredPath, displayName)
            };
        }

        return BuildClientQuarterLayout(basePath, displayName, scanDate);
    }

    public static string SanitizeClientFolderName(string companyName)
    {
        var name = string.IsNullOrWhiteSpace(companyName) ? "Client" : companyName.Trim();
        var windowsInvalid = new[] { ':', '*', '?', '"', '<', '>', '|' };
        foreach (var invalid in Path.GetInvalidFileNameChars().Concat(windowsInvalid.Except(Path.GetInvalidFileNameChars())))
            name = name.Replace(invalid, '_');

        name = name.Trim().TrimEnd('.');
        return string.IsNullOrWhiteSpace(name) ? "Client" : name;
    }

    private static ReportOutputLayout BuildClientQuarterLayout(string root, string companyName, string scanDate)
    {
        var clientPath = Path.Combine(root, SanitizeClientFolderName(companyName));
        var outputPath = BuildQuarterOutputPath(clientPath, scanDate);
        EnsureDirectory(outputPath);

        return new ReportOutputLayout
        {
            OutputDirectory = outputPath,
            TextOutputDirectory = outputPath,
            UsesStructuredPaths = true,
            UsesMiscSubfolder = false,
            ReportsPathPartial = GetReportsPathPartial(outputPath, companyName)
        };
    }

    public static string GetReportTimestamp() => DateTime.Now.ToString("yyyy-MM-dd_HHmmss");

    public static string GetSafeReportOutputPath(string targetDir, string companyName, string reportSuffix, string ext)
    {
        var name = string.IsNullOrWhiteSpace(companyName) ? "Client" : companyName.Trim();
        var filename = $"{name}{reportSuffix}.{ext}";
        var fullPath = Path.Combine(targetDir, filename);
        if (fullPath.Length <= MaxPathLength)
            return fullPath;

        var suffixPart = $"{reportSuffix}.{ext}";
        var availableForName = MaxPathLength - targetDir.Length - 1 - suffixPart.Length;
        var safeName = availableForName < name.Length
            ? availableForName > 5
                ? name[..availableForName]
                : name[..Math.Min(5, name.Length)]
            : name;

        return Path.Combine(targetDir, $"{safeName}{suffixPart}");
    }

    public static string? GetReportsPathPartial(string fullOutputPath, string? companyName = null)
    {
        if (string.IsNullOrWhiteSpace(fullOutputPath))
            return null;

        var path = Path.GetFullPath(fullOutputPath.Trim());
        var ndIdx = path.IndexOf("Network Documentation", StringComparison.OrdinalIgnoreCase);
        string? partial;
        if (ndIdx >= 0)
        {
            partial = path[ndIdx..].Replace(Path.DirectorySeparatorChar, '\\');
            if (Path.AltDirectorySeparatorChar != Path.DirectorySeparatorChar)
                partial = partial.Replace(Path.AltDirectorySeparatorChar, '\\');
        }
        else
        {
            var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
            partial = parts.Length >= 1 ? parts[^1] : null;
        }

        if (string.IsNullOrWhiteSpace(partial))
            return null;

        if (!string.IsNullOrWhiteSpace(companyName))
            return $"{companyName.Trim()}\\{partial}";

        return partial;
    }

    public static string ResolveVulnerabilityScansSubpath(string folderName)
    {
        if (string.IsNullOrWhiteSpace(folderName))
            return folderName;

        if (folderName.Contains("Vulnerability Scans", StringComparison.OrdinalIgnoreCase))
            return folderName;

        return Path.Combine(folderName, VulnerabilityScansSegment);
    }

    /// <summary>
    /// Picks a quarter folder under the client vulnerability-scans path.
    /// Uses the bare quarter (e.g. 2026 - Q2) when free; otherwise adds scan date, then a time suffix
    /// so multiple reviews in the same quarter/day do not overwrite the same folder.
    /// </summary>
    public static string ResolveQuarterFolderName(string clientPath, string scanDate)
    {
        var baseQuarter = GetQuarterFromDate(scanDate);
        var dateStr = TryGetScanDateString(scanDate, out var parsed)
            ? parsed.ToString("yyyy-MM-dd")
            : DateTime.Now.ToString("yyyy-MM-dd");

        var candidates = new List<string> { baseQuarter, $"{baseQuarter} {dateStr}" };
        var stamp = DateTime.Now.ToString("HHmmss");
        candidates.Add($"{baseQuarter} {dateStr}_{stamp}");

        foreach (var name in candidates)
        {
            if (!Directory.Exists(Path.Combine(clientPath, name)))
                return name;
        }

        return $"{baseQuarter} {dateStr}_{DateTime.Now:yyyyMMdd_HHmmss}";
    }

    private static bool TryGetScanDateString(string scanDate, out DateTime parsed)
    {
        parsed = default;
        if (string.IsNullOrWhiteSpace(scanDate))
            return false;

        return DateTime.TryParse(scanDate, out parsed);
    }

    public static string GetQuarterFromDate(string scanDate)
    {
        if (string.IsNullOrWhiteSpace(scanDate))
        {
            var now = DateTime.Now;
            return $"{now.Year} - Q{(int)Math.Ceiling(now.Month / 3.0)}";
        }

        if (DateTime.TryParse(scanDate, out var date))
            return $"{date.Year} - Q{(int)Math.Ceiling(date.Month / 3.0)}";

        var fallback = DateTime.Now;
        return $"{fallback.Year} - Q{(int)Math.Ceiling(fallback.Month / 3.0)}";
    }

    private bool TryResolveStructuredPath(string basePath, int companyId, string companyName, string scanDate, out string outputPath)
    {
        outputPath = "";

        if (companyId > 0 && companyFolderMapService.TryGetFolder(companyId, out var mappedFolder))
        {
            var resolvedFolder = ResolveVulnerabilityScansSubpath(mappedFolder);
            if (!string.Equals(mappedFolder, resolvedFolder, StringComparison.OrdinalIgnoreCase))
                companyFolderMapService.SetFolder(companyId, resolvedFolder);

            var clientPath = Path.Combine(basePath, resolvedFolder);
            outputPath = BuildQuarterOutputPath(clientPath, scanDate);
            return true;
        }

        var match = FindBestMatchingFolder(basePath, companyName);
        if (!string.IsNullOrWhiteSpace(match))
        {
            var folderName = ResolveVulnerabilityScansSubpath(match);
            if (companyId > 0)
                companyFolderMapService.SetFolder(companyId, folderName);

            var clientPath = Path.Combine(basePath, folderName);
            outputPath = BuildQuarterOutputPath(clientPath, scanDate);
            return true;
        }

        return false;
    }

    private static string BuildQuarterOutputPath(string clientPath, string scanDate)
    {
        var quarterFolder = ResolveQuarterFolderName(clientPath, scanDate);
        return Path.Combine(clientPath, quarterFolder);
    }

    private static string ResolveFallbackDirectory(UserSettings settings, string? fallbackPath)
    {
        if (!string.IsNullOrWhiteSpace(fallbackPath) && Directory.Exists(fallbackPath))
            return Path.GetFullPath(fallbackPath.Trim());

        if (!string.IsNullOrWhiteSpace(settings.LastOutputDirectory) && Directory.Exists(settings.LastOutputDirectory))
            return Path.GetFullPath(settings.LastOutputDirectory.Trim());

        return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    }

    private static void EnsureDirectory(string path) =>
        Directory.CreateDirectory(path);

    private static string? FindBestMatchingFolder(string basePath, string companyName)
    {
        string[] subfolders;
        try
        {
            subfolders = Directory.GetDirectories(basePath).Select(Path.GetFileName).Where(n => n is not null).Cast<string>().ToArray();
        }
        catch
        {
            return null;
        }

        var normLower = NormalizeCompanyName(companyName);

        string? bestMatch = null;
        var bestScore = -1;

        foreach (var folder in subfolders)
        {
            if (string.Equals(folder, "Misc", StringComparison.OrdinalIgnoreCase))
                continue;

            var folderNormLower = NormalizeCompanyName(folder);
            var fLower = folder.ToLowerInvariant();
            var score = 0;
            if (normLower == folderNormLower)
                score = 100;
            else if (normLower.Length > 0 && (fLower.Contains(normLower, StringComparison.Ordinal) || normLower.Contains(fLower, StringComparison.Ordinal)))
                score = 50;
            else if (normLower.Length > 0 && folderNormLower.Length > 0 &&
                     (folderNormLower.Contains(normLower, StringComparison.Ordinal) || normLower.Contains(folderNormLower, StringComparison.Ordinal)))
                score = 50;
            else if (normLower.Length >= 3 && fLower.Length >= 3)
            {
                var nPre = normLower[..Math.Min(5, normLower.Length)];
                var fPre = fLower[..Math.Min(5, fLower.Length)];
                if (fLower.StartsWith(nPre, StringComparison.Ordinal) || normLower.StartsWith(fPre, StringComparison.Ordinal))
                    score = 25;
            }

            if (score > bestScore)
            {
                bestScore = score;
                bestMatch = folder;
            }
            else if (score == bestScore && bestMatch is not null)
            {
                bestMatch = null;
            }
        }

        return bestScore > 0 ? bestMatch : null;
    }

    internal static string NormalizeCompanyName(string companyName)
    {
        var norm = Regex.Replace(companyName.Trim(), @"\s*(Inc|LLC|Corp|ltd)\.?\s*$", "", RegexOptions.IgnoreCase).Trim();
        norm = Regex.Replace(norm, @"\s+", " ");
        return norm.ToLowerInvariant();
    }
}
