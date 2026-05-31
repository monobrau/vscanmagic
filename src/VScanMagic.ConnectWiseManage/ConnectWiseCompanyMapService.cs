using System.Text.Json;
using VScanMagic.Core.Paths;

namespace VScanMagic.ConnectWiseManage;

public sealed class ConnectWiseCompanyMapService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, PropertyNameCaseInsensitive = true };
    private Dictionary<string, ConnectWiseCompanyMapEntry> _map = new(StringComparer.Ordinal);
    private bool _loaded;

    public IReadOnlyDictionary<string, ConnectWiseCompanyMapEntry> Map
    {
        get
        {
            EnsureLoaded();
            return _map;
        }
    }

    public bool TryGetManageCompanyId(string connectSecureCompanyId, out int manageCompanyId, out string? manageCompanyName)
    {
        EnsureLoaded();
        if (_map.TryGetValue(connectSecureCompanyId.Trim(), out var entry))
        {
            manageCompanyId = entry.ManageCompanyId;
            manageCompanyName = entry.ManageCompanyName;
            return entry.ManageCompanyId > 0;
        }

        manageCompanyId = 0;
        manageCompanyName = null;
        return false;
    }

    public void ReplaceAll(IEnumerable<KeyValuePair<string, ConnectWiseCompanyMapEntry>> entries)
    {
        _loaded = true;
        _map = entries.ToDictionary(
            e => e.Key.Trim(),
            e => e.Value,
            StringComparer.Ordinal);
        Save();
    }

    public IReadOnlyList<KeyValuePair<string, ConnectWiseCompanyMapEntry>> GetOrderedEntries()
    {
        EnsureLoaded();
        return _map.OrderBy(p => p.Value.ManageCompanyName, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private void EnsureLoaded()
    {
        if (_loaded)
            return;

        _loaded = true;
        var path = VScanMagicPaths.ConnectWiseCompanyMapFile();
        if (!File.Exists(path))
            return;

        try
        {
            var json = File.ReadAllText(path);
            var map = JsonSerializer.Deserialize<Dictionary<string, ConnectWiseCompanyMapEntry>>(json, JsonOptions);
            if (map is not null)
                _map = new Dictionary<string, ConnectWiseCompanyMapEntry>(map, StringComparer.Ordinal);
        }
        catch
        {
            _map = new Dictionary<string, ConnectWiseCompanyMapEntry>(StringComparer.Ordinal);
        }
    }

    private void Save()
    {
        var path = VScanMagicPaths.ConnectWiseCompanyMapFile();
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path, JsonSerializer.Serialize(_map, JsonOptions));
    }
}
