using System.Text.Json;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class RmitPlusSettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private Dictionary<string, bool> _settings = new(StringComparer.OrdinalIgnoreCase);
    private bool _loaded;

    public bool TryGet(string companyName, out bool isRmitPlus)
    {
        EnsureLoaded();
        return _settings.TryGetValue(NormalizeKey(companyName), out isRmitPlus);
    }

    public bool GetOrDefault(string companyName, bool defaultValue = false)
    {
        EnsureLoaded();
        return _settings.TryGetValue(NormalizeKey(companyName), out var value) ? value : defaultValue;
    }

    public void Set(string companyName, bool isRmitPlus)
    {
        if (string.IsNullOrWhiteSpace(companyName))
            return;

        EnsureLoaded();
        _settings[NormalizeKey(companyName)] = isRmitPlus;
        Save();
    }

    private void EnsureLoaded()
    {
        if (_loaded)
            return;

        _loaded = true;
        var path = VScanMagicPaths.RmitPlusSettingsFile();
        if (!File.Exists(path))
            return;

        try
        {
            var json = File.ReadAllText(path);
            var map = JsonSerializer.Deserialize<Dictionary<string, bool>>(json, JsonOptions);
            if (map is not null)
                _settings = new Dictionary<string, bool>(map, StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            _settings = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void Save()
    {
        var path = VScanMagicPaths.RmitPlusSettingsFile();
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path, JsonSerializer.Serialize(_settings, JsonOptions));
    }

    private static string NormalizeKey(string companyName) => companyName.Trim();
}
