using System.Text.Json;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed record PatchActivityEntry(
    int CompanyId,
    string JobId,
    string Type,
    string Status,
    string Description,
    string? HostName,
    string? AgentIp,
    DateTimeOffset RequestedAt,
    string? ConnectSecureMessage);

public sealed class PatchActivityHistoryService(string? configDirectory = null)
{
    private const int MaxEntries = 200;
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public IReadOnlyList<PatchActivityEntry> GetEntries(int companyId, int limit = 50)
    {
        if (companyId <= 0 || limit <= 0)
            return [];

        return LoadDocument()
            .Entries
            .Where(entry => entry.CompanyId == companyId)
            .OrderByDescending(entry => entry.RequestedAt)
            .Take(limit)
            .ToList();
    }

    public void Record(PatchActivityEntry entry)
    {
        if (entry.CompanyId <= 0)
            return;

        var document = LoadDocument();
        var existing = document.Entries
            .Where(e => !string.Equals(e.JobId, entry.JobId, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(e => e.RequestedAt)
            .Take(MaxEntries - 1)
            .ToList();

        existing.Insert(0, entry);
        document.Entries = existing;
        SaveDocument(document);
    }

    private PatchActivityHistoryDocument LoadDocument()
    {
        var path = VScanMagicPaths.PatchActivityHistoryFile(configDirectory);
        if (!File.Exists(path))
            return new PatchActivityHistoryDocument();

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<PatchActivityHistoryDocument>(json, JsonOptions)
                   ?? new PatchActivityHistoryDocument();
        }
        catch
        {
            return new PatchActivityHistoryDocument();
        }
    }

    private void SaveDocument(PatchActivityHistoryDocument document)
    {
        var path = VScanMagicPaths.PatchActivityHistoryFile(configDirectory);
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        File.WriteAllText(path, JsonSerializer.Serialize(document, JsonOptions));
    }

    private sealed class PatchActivityHistoryDocument
    {
        public List<PatchActivityEntry> Entries { get; set; } = [];
    }
}
