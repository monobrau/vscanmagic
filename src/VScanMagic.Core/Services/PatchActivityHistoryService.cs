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
    string? ConnectSecureMessage,
    IReadOnlyList<int>? AgentIds = null,
    IReadOnlyList<int>? SolutionIds = null,
    string? Product = null,
    string? TargetFix = null,
    bool IsOsPatch = false,
    bool IsEndOfLife = false,
    IReadOnlyList<int>? OsAssetIds = null,
    string? VerificationSummary = null,
    DateTimeOffset? VerifiedAt = null);

public sealed class PatchActivityHistoryService(string? configDirectory = null)
{
    private const int MaxEntriesPerCompany = 100;
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly object _lock = new();

    public IReadOnlyList<PatchActivityEntry> GetEntries(int companyId, int limit = 50)
    {
        if (companyId <= 0 || limit <= 0)
            return [];

        lock (_lock)
        {
            return LoadDocumentUnsafe()
                .Entries
                .Where(entry => entry.CompanyId == companyId)
                .OrderByDescending(entry => entry.RequestedAt)
                .Take(limit)
                .ToList();
        }
    }

    public PatchActivityEntry? GetByJobId(int companyId, string jobId)
    {
        if (companyId <= 0 || string.IsNullOrWhiteSpace(jobId))
            return null;

        lock (_lock)
        {
            return LoadDocumentUnsafe()
                .Entries
                .FirstOrDefault(entry =>
                    entry.CompanyId == companyId &&
                    string.Equals(entry.JobId, jobId, StringComparison.OrdinalIgnoreCase));
        }
    }

    public bool UpdateEntry(PatchActivityEntry updated)
    {
        if (updated.CompanyId <= 0 || string.IsNullOrWhiteSpace(updated.JobId))
            return false;

        lock (_lock)
        {
            var document = LoadDocumentUnsafe();
            var index = document.Entries.FindIndex(entry =>
                entry.CompanyId == updated.CompanyId &&
                string.Equals(entry.JobId, updated.JobId, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                return false;

            document.Entries[index] = updated;
            SaveDocumentUnsafe(document);
            return true;
        }
    }

    public void Record(PatchActivityEntry entry)
    {
        if (entry.CompanyId <= 0)
            return;

        lock (_lock)
        {
            var document = LoadDocumentUnsafe();
            var otherCompanies = document.Entries
                .Where(e => e.CompanyId != entry.CompanyId)
                .ToList();
            var companyEntries = document.Entries
                .Where(e => e.CompanyId == entry.CompanyId &&
                            !string.Equals(e.JobId, entry.JobId, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(e => e.RequestedAt)
                .Take(MaxEntriesPerCompany - 1)
                .ToList();

            companyEntries.Insert(0, entry);
            document.Entries = otherCompanies.Concat(companyEntries).ToList();
            SaveDocumentUnsafe(document);
        }
    }

    private PatchActivityHistoryDocument LoadDocumentUnsafe()
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

    private void SaveDocumentUnsafe(PatchActivityHistoryDocument document)
    {
        var path = VScanMagicPaths.PatchActivityHistoryFile(configDirectory);
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var tempPath = path + ".tmp";
        File.WriteAllText(tempPath, JsonSerializer.Serialize(document, JsonOptions));
        File.Move(tempPath, path, overwrite: true);
    }

    private sealed class PatchActivityHistoryDocument
    {
        public List<PatchActivityEntry> Entries { get; set; } = [];
    }
}
