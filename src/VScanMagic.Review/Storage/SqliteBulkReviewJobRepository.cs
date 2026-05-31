using System.Text.Json;
using Microsoft.Data.Sqlite;
using VScanMagic.Core.Paths;
using VScanMagic.Review.Models;

namespace VScanMagic.Review.Storage;

public interface IBulkReviewJobRepository
{
    Task<IReadOnlyList<BulkReviewJob>> ListAsync(int limit = 20, CancellationToken ct = default);
    Task<BulkReviewJob?> GetAsync(string id, CancellationToken ct = default);
    Task SaveAsync(BulkReviewJob job, CancellationToken ct = default);
    Task RecoverInterruptedJobsAsync(CancellationToken ct = default);
}

public sealed class SqliteBulkReviewJobRepository : IBulkReviewJobRepository, IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false, PropertyNameCaseInsensitive = true };
    private readonly string _connectionString;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public SqliteBulkReviewJobRepository(string? configDirectory = null)
    {
        var sessionsDir = VScanMagicPaths.SessionsDirectory(configDirectory);
        Directory.CreateDirectory(sessionsDir);
        var dbPath = Path.Combine(sessionsDir, "bulk_jobs.db");
        _connectionString = new SqliteConnectionStringBuilder { DataSource = dbPath }.ConnectionString;
        EnsureSchema();
    }

    private void EnsureSchema()
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS bulk_review_jobs (
                id TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                payload_json TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_bulk_review_jobs_updated ON bulk_review_jobs(updated_at DESC);
            """;
        cmd.ExecuteNonQuery();
    }

    public async Task<IReadOnlyList<BulkReviewJob>> ListAsync(int limit = 20, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var results = new List<BulkReviewJob>();
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT payload_json FROM bulk_review_jobs ORDER BY updated_at DESC LIMIT $limit";
            cmd.Parameters.AddWithValue("$limit", limit);
            await using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                var job = JsonSerializer.Deserialize<BulkReviewJob>(reader.GetString(0), JsonOptions);
                if (job is not null)
                    results.Add(job);
            }

            return results;
        }
        finally { _lock.Release(); }
    }

    public async Task<BulkReviewJob?> GetAsync(string id, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT payload_json FROM bulk_review_jobs WHERE id = $id";
            cmd.Parameters.AddWithValue("$id", id);
            var result = await cmd.ExecuteScalarAsync(ct);
            return result is string json
                ? JsonSerializer.Deserialize<BulkReviewJob>(json, JsonOptions)
                : null;
        }
        finally { _lock.Release(); }
    }

    public async Task SaveAsync(BulkReviewJob job, CancellationToken ct = default)
    {
        job.UpdatedAt = DateTimeOffset.Now;
        var json = JsonSerializer.Serialize(job, JsonOptions);
        await SaveJsonAsync(job, json, ct);
    }

    public async Task RecoverInterruptedJobsAsync(CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var interrupted = new List<BulkReviewJob>();
            await using (var conn = new SqliteConnection(_connectionString))
            {
                await conn.OpenAsync(ct);
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    SELECT payload_json FROM bulk_review_jobs
                    WHERE status IN ($running, $queued)
                    """;
                cmd.Parameters.AddWithValue("$running", BulkReviewJobStatus.Running.ToString());
                cmd.Parameters.AddWithValue("$queued", BulkReviewJobStatus.Queued.ToString());
                await using var reader = await cmd.ExecuteReaderAsync(ct);
                while (await reader.ReadAsync(ct))
                {
                    var job = JsonSerializer.Deserialize<BulkReviewJob>(reader.GetString(0), JsonOptions);
                    if (job is not null)
                        interrupted.Add(job);
                }
            }

            foreach (var job in interrupted)
            {
                job.Status = BulkReviewJobStatus.Failed;
                job.ErrorMessage = "Interrupted by server restart.";
                foreach (var item in job.Items)
                {
                    if (item.Phase is BulkReviewItemPhase.Completed or BulkReviewItemPhase.Skipped)
                        continue;

                    item.Phase = BulkReviewItemPhase.Failed;
                    item.StatusMessage = "Failed.";
                    item.ErrorMessage ??= "Interrupted by server restart.";
                }

                job.UpdatedAt = DateTimeOffset.Now;
                var json = JsonSerializer.Serialize(job, JsonOptions);
                await SaveJsonUnlockedAsync(job, json, ct);
            }
        }
        finally { _lock.Release(); }
    }

    private async Task SaveJsonAsync(BulkReviewJob job, string json, CancellationToken ct)
    {
        await _lock.WaitAsync(ct);
        try
        {
            await SaveJsonUnlockedAsync(job, json, ct);
        }
        finally { _lock.Release(); }
    }

    private async Task SaveJsonUnlockedAsync(BulkReviewJob job, string json, CancellationToken ct)
    {
        await using var conn = new SqliteConnection(_connectionString);
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO bulk_review_jobs (id, status, created_at, updated_at, payload_json)
            VALUES ($id, $status, $created, $updated, $json)
            ON CONFLICT(id) DO UPDATE SET
                status = $status, updated_at = $updated, payload_json = $json
            """;
        cmd.Parameters.AddWithValue("$id", job.Id);
        cmd.Parameters.AddWithValue("$status", job.Status.ToString());
        cmd.Parameters.AddWithValue("$created", job.CreatedAt.ToString("O"));
        cmd.Parameters.AddWithValue("$updated", job.UpdatedAt.ToString("O"));
        cmd.Parameters.AddWithValue("$json", json);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    public void Dispose() => _lock.Dispose();
}
