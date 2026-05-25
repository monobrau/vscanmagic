using System.Text.Json;
using Microsoft.Data.Sqlite;
using VScanMagic.Core.Paths;
using VScanMagic.Review.Models;

namespace VScanMagic.Review.Storage;

public interface IReviewSessionRepository
{
    Task<IReadOnlyList<ReviewSession>> ListAsync(bool includeArchived = false, CancellationToken ct = default);
    Task<ReviewSession?> GetAsync(string id, CancellationToken ct = default);
    Task SaveAsync(ReviewSession session, CancellationToken ct = default);
    Task DeleteAsync(string id, CancellationToken ct = default);
}

public sealed class SqliteReviewSessionRepository : IReviewSessionRepository, IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false, PropertyNameCaseInsensitive = true };
    private readonly string _connectionString;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public SqliteReviewSessionRepository(string? configDirectory = null)
    {
        var sessionsDir = VScanMagicPaths.SessionsDirectory(configDirectory);
        Directory.CreateDirectory(sessionsDir);
        var dbPath = Path.Combine(sessionsDir, "reviews.db");
        _connectionString = new SqliteConnectionStringBuilder { DataSource = dbPath }.ConnectionString;
        EnsureSchema();
    }

    private void EnsureSchema()
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS review_sessions (
                id TEXT PRIMARY KEY,
                client_name TEXT NOT NULL,
                scan_date TEXT NOT NULL,
                company_id TEXT,
                presenter TEXT,
                source_file_path TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                payload_json TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_review_sessions_updated ON review_sessions(updated_at DESC);
            """;
        cmd.ExecuteNonQuery();
    }

    public async Task<IReadOnlyList<ReviewSession>> ListAsync(bool includeArchived = false, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var results = new List<ReviewSession>();
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT payload_json FROM review_sessions ORDER BY updated_at DESC";
            await using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                var json = reader.GetString(0);
                var session = JsonSerializer.Deserialize<ReviewSession>(json, JsonOptions);
                if (session is null)
                    continue;

                session.Findings ??= [];
                if (!includeArchived && session.IsArchived)
                    continue;

                results.Add(session);
            }
            return results;
        }
        finally { _lock.Release(); }
    }

    public async Task<ReviewSession?> GetAsync(string id, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT payload_json FROM review_sessions WHERE id = $id";
            cmd.Parameters.AddWithValue("$id", id);
            var result = await cmd.ExecuteScalarAsync(ct);
            if (result is not string json) return null;
            var session = JsonSerializer.Deserialize<ReviewSession>(json, JsonOptions);
            if (session is not null)
                session.Findings ??= [];
            return session;
        }
        finally { _lock.Release(); }
    }

    public async Task SaveAsync(ReviewSession session, CancellationToken ct = default)
    {
        session.UpdatedAt = DateTimeOffset.UtcNow;
        var json = JsonSerializer.Serialize(session, JsonOptions);

        await _lock.WaitAsync(ct);
        try
        {
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO review_sessions (id, client_name, scan_date, company_id, presenter, source_file_path, created_at, updated_at, payload_json)
                VALUES ($id, $client, $scan, $company, $presenter, $source, $created, $updated, $json)
                ON CONFLICT(id) DO UPDATE SET
                    client_name = $client, scan_date = $scan, company_id = $company, presenter = $presenter,
                    source_file_path = $source, updated_at = $updated, payload_json = $json
                """;
            cmd.Parameters.AddWithValue("$id", session.Id);
            cmd.Parameters.AddWithValue("$client", session.ClientName);
            cmd.Parameters.AddWithValue("$scan", session.ScanDate);
            cmd.Parameters.AddWithValue("$company", (object?)session.CompanyId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$presenter", session.Presenter);
            cmd.Parameters.AddWithValue("$source", (object?)session.SourceFilePath ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$created", session.CreatedAt.ToString("O"));
            cmd.Parameters.AddWithValue("$updated", session.UpdatedAt.ToString("O"));
            cmd.Parameters.AddWithValue("$json", json);
            await cmd.ExecuteNonQueryAsync(ct);
        }
        finally { _lock.Release(); }
    }

    public async Task DeleteAsync(string id, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            await using var conn = new SqliteConnection(_connectionString);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM review_sessions WHERE id = $id";
            cmd.Parameters.AddWithValue("$id", id);
            await cmd.ExecuteNonQueryAsync(ct);
        }
        finally { _lock.Release(); }
    }

    public void Dispose() => _lock.Dispose();
}
