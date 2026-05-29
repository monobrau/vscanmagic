using System.Text.RegularExpressions;
using VScanMagic.Core.Models;
using VScanMagic.Core.Services;

namespace VScanMagic.Core.Paths;

public sealed class ReportPathResolver(CompanyFolderMapService companyFolderMapService)
{
    private const int MaxPathLength = 250;
    public const string MiscSubfolderName = "Misc";
    private const string PendingEpssReportType = "pending-epss";
    private const string VulnerabilityScansSegment = "Network Documentation/Vulnerability Scans";
    private static readonly Regex QuarterFolderPattern = new(
        @"^\d{4}\s*-\s*Q[1-4](\b|\s)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public ReportOutputLayout Resolve(UserSettings settings, int companyId, string companyName, string scanDate, string? fallbackPath = null)
    {
        var configuredBase = NormalizeConfiguredBasePath(settings.ReportsBasePath);
        var usesStructuredBase = !string.IsNullOrWhiteSpace(configuredBase);
        // When Reports Base Path is set, always resolve from that root — never from a prior quarter output folder.
        var fallback = ResolveFallbackDirectory(settings, fallbackPath, configuredBase);
        var displayName = string.IsNullOrWhiteSpace(companyName) ? "Client" : companyName.Trim();

        var basePath = configuredBase;
        if (string.IsNullOrWhiteSpace(basePath) || !Directory.Exists(basePath))
        {
            if (usesStructuredBase)
            {
                var root = Path.GetFullPath(basePath!);
                return BuildClientQuarterLayout(root, displayName, scanDate);
            }

            EnsureDirectory(fallback);
            return BuildFlatLayout(fallback);
        }

        basePath = Path.GetFullPath(basePath);

        if (TryResolveStructuredPath(basePath, companyId, displayName, scanDate, out var structuredPath))
        {
            return BuildStructuredLayout(structuredPath, displayName);
        }

        return BuildClientQuarterLayout(basePath, displayName, scanDate);
    }

    public static string GetDefaultManualOutputDirectory(UserSettings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.LastOutputDirectory))
            return settings.LastOutputDirectory.Trim();

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "VScanMagic",
            "Exports");
    }

    public static string GetDownloadDirectory(ReportOutputLayout layout, string? reportType = null)
    {
        if (layout.UsesMiscSubfolder &&
            string.Equals(reportType, PendingEpssReportType, StringComparison.OrdinalIgnoreCase))
        {
            EnsureDirectory(layout.TextOutputDirectory);
            return layout.TextOutputDirectory;
        }

        EnsureDirectory(layout.OutputDirectory);
        return layout.OutputDirectory;
    }

    /// <summary>Quarter folder for Top N Word report and ConnectSecure downloads.</summary>
    public static string GetTopNReportDirectory(ReportOutputLayout layout)
    {
        EnsureDirectory(layout.OutputDirectory);
        return layout.OutputDirectory;
    }

    /// <summary>Misc folder for supplemental exports (PDF review, data XLSX, host counts) when structured paths are enabled.</summary>
    public static string GetSupplementalExportDirectory(ReportOutputLayout layout)
    {
        if (layout.UsesMiscSubfolder)
        {
            EnsureDirectory(layout.TextOutputDirectory);
            return layout.TextOutputDirectory;
        }

        EnsureDirectory(layout.OutputDirectory);
        return layout.OutputDirectory;
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
        return BuildStructuredLayout(outputPath, companyName);
    }

    private static ReportOutputLayout BuildStructuredLayout(string outputPath, string companyName)
    {
        EnsureDirectory(outputPath);
        var miscPath = Path.Combine(outputPath, MiscSubfolderName);
        EnsureDirectory(miscPath);

        return new ReportOutputLayout
        {
            OutputDirectory = outputPath,
            TextOutputDirectory = miscPath,
            UsesStructuredPaths = true,
            UsesMiscSubfolder = true,
            ReportsPathPartial = GetReportsPathPartial(outputPath, companyName)
        };
    }

    private static ReportOutputLayout BuildFlatLayout(string outputPath)
    {
        EnsureDirectory(outputPath);
        return new ReportOutputLayout
        {
            OutputDirectory = outputPath,
            TextOutputDirectory = outputPath,
            UsesStructuredPaths = false,
            UsesMiscSubfolder = false,
            ReportsPathPartial = null
        };
    }

    /// <summary>Rebuild layout metadata for an existing on-disk quarter folder (e.g. reuse after ConnectSecure downloads).</summary>
    public static ReportOutputLayout LayoutForExistingDirectory(string outputDirectory, string companyName)
    {
        var full = Path.GetFullPath(outputDirectory.Trim());
        var misc = Path.Combine(full, MiscSubfolderName);
        var usesMisc = Directory.Exists(misc);
        if (usesMisc)
            EnsureDirectory(misc);

        EnsureDirectory(full);
        return new ReportOutputLayout
        {
            OutputDirectory = full,
            TextOutputDirectory = usesMisc ? misc : full,
            UsesStructuredPaths = true,
            UsesMiscSubfolder = usesMisc,
            ReportsPathPartial = GetReportsPathPartial(full, companyName)
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
    /// Single quarter folder per client (e.g. 2026 - Q2). ConnectSecure downloads and Top N Word go here;
    /// supplemental VScanMagic exports use the Misc subfolder.
    /// </summary>
    public static string ResolveQuarterFolderName(string clientPath, string scanDate) =>
        GetQuarterFromDate(scanDate);

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
            mappedFolder = SanitizeMappedFolderPath(mappedFolder);
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

    private static string ResolveFallbackDirectory(UserSettings settings, string? fallbackPath, string? configuredBasePath)
    {
        if (!string.IsNullOrWhiteSpace(configuredBasePath))
            return Path.GetFullPath(configuredBasePath.Trim());

        if (!string.IsNullOrWhiteSpace(fallbackPath) && Directory.Exists(fallbackPath))
            return Path.GetFullPath(fallbackPath.Trim());

        if (!string.IsNullOrWhiteSpace(settings.LastOutputDirectory) && Directory.Exists(settings.LastOutputDirectory))
            return Path.GetFullPath(settings.LastOutputDirectory.Trim());

        return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    }

    /// <summary>
    /// Strips accidental quarter/timestamp segments from stored folder mappings so paths do not nest on each resolve.
    /// </summary>
    /// <summary>
    /// Strips quarter folders and accidental duplicate tail segments from stored company folder mappings.
    /// </summary>
    public static string SanitizeMappedFolderPath(string mappedFolder)
    {
        if (string.IsNullOrWhiteSpace(mappedFolder))
            return mappedFolder;

        var parts = mappedFolder
            .Replace('\\', Path.DirectorySeparatorChar)
            .Replace('/', Path.DirectorySeparatorChar)
            .Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        var kept = new List<string>();
        foreach (var part in parts)
        {
            if (IsQuarterOrDatedFolder(part))
                continue;

            kept.Add(part);
            if (part.Equals("Vulnerability Scans", StringComparison.OrdinalIgnoreCase))
                break;
        }

        if (kept.Count == 0)
        {
            var firstClient = parts.FirstOrDefault(p => !IsQuarterOrDatedFolder(p));
            if (!string.IsNullOrWhiteSpace(firstClient))
                kept.Add(firstClient);
        }

        return kept.Count > 0 ? string.Join(Path.DirectorySeparatorChar, kept) : mappedFolder.Trim();
    }

    /// <summary>
    /// Ensures Reports Base Path points at the tenant root (e.g. ...\Global), not a quarter or review output folder.
    /// </summary>
    public static string NormalizeConfiguredBasePath(string? configuredBasePath)
    {
        if (string.IsNullOrWhiteSpace(configuredBasePath))
            return "";

        var full = Path.GetFullPath(configuredBasePath.Trim());
        var root = Path.GetPathRoot(full) ?? "";
        var relative = full[root.Length..].TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (string.IsNullOrWhiteSpace(relative))
            return full;

        var parts = relative.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        var kept = parts.Where(p => !IsQuarterOrDatedFolder(p)).ToList();
        if (kept.Count == 0)
            return full;

        var globalIdx = kept.FindIndex(p => p.Equals("Global", StringComparison.OrdinalIgnoreCase));
        if (globalIdx >= 0)
            kept = kept.Take(globalIdx + 1).ToList();

        return string.IsNullOrEmpty(root)
            ? string.Join(Path.DirectorySeparatorChar, kept)
            : Path.Combine(root, string.Join(Path.DirectorySeparatorChar, kept));
    }

    public static bool IsQuarterOrDatedFolder(string folderName) =>
        !string.IsNullOrWhiteSpace(folderName) && QuarterFolderPattern.IsMatch(folderName.Trim());

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

        var companyIsGlobal = string.Equals(companyName, "Global", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(companyName, "All Companies", StringComparison.OrdinalIgnoreCase);

        foreach (var folder in subfolders)
        {
            if (string.Equals(folder, "Misc", StringComparison.OrdinalIgnoreCase))
                continue;

            if (IsQuarterOrDatedFolder(folder))
                continue;

            if (!companyIsGlobal && string.Equals(folder, "Global", StringComparison.OrdinalIgnoreCase))
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
