using System.Text.Json;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class CompanyFolderMapService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private Dictionary<string, string> _map = new(StringComparer.Ordinal);
    private bool _loaded;

    public IReadOnlyDictionary<string, string> Map
    {
        get
        {
            EnsureLoaded();
            return _map;
        }
    }

    public bool TryGetFolder(int companyId, out string folderName)
    {
        EnsureLoaded();
        if (_map.TryGetValue(companyId.ToString(), out folderName!))
        {
            folderName = NormalizeStoredPath(folderName);
            return true;
        }

        folderName = "";
        return false;
    }

    public void SetFolder(int companyId, string folderName)
    {
        EnsureLoaded();
        _map[companyId.ToString()] = NormalizeStoredPath(folderName);
        Save();
    }

    public void RemoveFolder(int companyId)
    {
        EnsureLoaded();
        if (_map.Remove(companyId.ToString()))
            Save();
    }

    public void ReplaceAll(IEnumerable<KeyValuePair<string, string>> entries)
    {
        _loaded = true;
        _map = entries.ToDictionary(
            e => e.Key.Trim(),
            e => NormalizeStoredPath(e.Value),
            StringComparer.Ordinal);
        Save();
    }

    public IReadOnlyList<KeyValuePair<string, string>> GetOrderedEntries()
    {
        EnsureLoaded();
        return _map.OrderBy(p => p.Value, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static string NormalizeStoredPath(string folderName) =>
        folderName.Replace('\\', Path.DirectorySeparatorChar)
            .Replace('/', Path.DirectorySeparatorChar)
            .Trim();

    public void EnsureLoaded()
    {
        if (_loaded)
            return;

        _loaded = true;
        var path = VScanMagicPaths.CompanyFolderMapFile();
        if (!File.Exists(path))
            return;

        try
        {
            var json = File.ReadAllText(path);
            var map = JsonSerializer.Deserialize<Dictionary<string, string>>(json, JsonOptions);
            if (map is not null)
            {
                _map = map.ToDictionary(
                    pair => pair.Key,
                    pair => NormalizeStoredPath(pair.Value),
                    StringComparer.Ordinal);
            }
        }
        catch
        {
            _map = new Dictionary<string, string>(StringComparer.Ordinal);
        }
    }

    private void Save()
    {
        var path = VScanMagicPaths.CompanyFolderMapFile();
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path, JsonSerializer.Serialize(_map, JsonOptions));
    }
}
