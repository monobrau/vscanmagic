using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed record ReportFolderHistoryEntry(string CompanyName, string OutputPath, string ProcessedAt);

public sealed class ReportFolderHistoryService(SettingsService settingsService)
{
    private const int MaxEntries = 100;
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public IReadOnlyList<ReportFolderHistoryEntry> GetEntries()
    {
        var path = VScanMagicPaths.ReportFolderHistoryFile();
        if (!File.Exists(path))
            return [];

        try
        {
            var json = File.ReadAllText(path);
            var doc = JsonSerializer.Deserialize<ReportFolderHistoryDocument>(json, JsonOptions);
            return doc?.Entries ?? [];
        }
        catch
        {
            return [];
        }
    }

    public void Add(string companyName, string outputPath)
    {
        if (string.IsNullOrWhiteSpace(outputPath) || !Directory.Exists(outputPath))
            return;

        var pathNorm = Path.GetFullPath(outputPath.Trim());
        var settings = settingsService.LoadUserSettings();
        var basePath = settings.ReportsBasePath?.Trim();
        if (!string.IsNullOrWhiteSpace(basePath) && Directory.Exists(basePath))
        {
            var baseNorm = Path.GetFullPath(basePath);
            if (!pathNorm.StartsWith(baseNorm, OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal))
            {
                return;
            }
        }

        var existing = GetEntries()
            .Where(e => !string.Equals(e.OutputPath, pathNorm, StringComparison.OrdinalIgnoreCase))
            .Take(MaxEntries - 1)
            .ToList();

        var entry = new ReportFolderHistoryEntry(
            companyName.Trim(),
            pathNorm,
            DateTime.Now.ToString("yyyy-MM-dd HH:mm"));

        var entries = new List<ReportFolderHistoryEntry> { entry };
        entries.AddRange(existing);

        var filePath = VScanMagicPaths.ReportFolderHistoryFile();
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        File.WriteAllText(filePath, JsonSerializer.Serialize(
            new ReportFolderHistoryDocument { Entries = entries },
            JsonOptions));
    }

    public string? GetLatestOutputPath(string companyName)
    {
        if (string.IsNullOrWhiteSpace(companyName))
            return null;

        return GetEntries()
            .FirstOrDefault(e => string.Equals(e.CompanyName, companyName.Trim(), StringComparison.OrdinalIgnoreCase))
            ?.OutputPath;
    }

    private sealed class ReportFolderHistoryDocument
    {
        public List<ReportFolderHistoryEntry> Entries { get; set; } = [];
    }
}
