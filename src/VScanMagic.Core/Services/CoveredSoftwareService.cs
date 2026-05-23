using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class CoveredSoftwareService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };
    private List<CoveredSoftwareEntry>? _cache;

    public IReadOnlyList<CoveredSoftwareEntry> Load()
    {
        if (_cache is not null)
            return _cache;

        var path = VScanMagicPaths.CoveredSoftwareFile();
        if (!File.Exists(path))
        {
            _cache = GetDefaults();
            Save(_cache);
            return _cache;
        }

        try
        {
            var json = File.ReadAllText(path);
            _cache = JsonSerializer.Deserialize<List<CoveredSoftwareEntry>>(json, JsonOptions) ?? GetDefaults();
        }
        catch
        {
            _cache = GetDefaults();
        }

        return _cache;
    }

    public void Save(IEnumerable<CoveredSoftwareEntry> entries)
    {
        _cache = entries.ToList();
        var path = VScanMagicPaths.CoveredSoftwareFile();
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path, JsonSerializer.Serialize(_cache, JsonOptions));
    }

    public static List<CoveredSoftwareEntry> GetDefaults() =>
    [
        new CoveredSoftwareEntry { Pattern = "*Microsoft*", IsPattern = true },
        new CoveredSoftwareEntry { Pattern = "*Adobe*", IsPattern = true },
        new CoveredSoftwareEntry { Pattern = "*Google Chrome*", IsPattern = true },
        new CoveredSoftwareEntry { Pattern = "*Mozilla Firefox*", IsPattern = true }
    ];
}
