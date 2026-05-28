using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class TemplatesService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private VScanMagicTemplates? _cache;

    public VScanMagicTemplates Load()
    {
        if (_cache is not null)
            return _cache;

        var path = VScanMagicPaths.TemplatesFile();
        if (!File.Exists(path))
        {
            _cache = new VScanMagicTemplates();
            return _cache;
        }

        try
        {
            var json = File.ReadAllText(path);
            _cache = JsonSerializer.Deserialize<VScanMagicTemplates>(json, JsonOptions) ?? new VScanMagicTemplates();
        }
        catch
        {
            _cache = new VScanMagicTemplates();
        }

        return _cache;
    }

    public void Save(VScanMagicTemplates templates)
    {
        var dir = VScanMagicPaths.GetConfigDirectory();
        Directory.CreateDirectory(dir);
        _cache = templates;
        File.WriteAllText(VScanMagicPaths.TemplatesFile(),
            JsonSerializer.Serialize(templates, new JsonSerializerOptions { WriteIndented = true }));
    }
}
