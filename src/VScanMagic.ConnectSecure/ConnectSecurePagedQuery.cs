using System.Diagnostics;
using System.Text.Json;

namespace VScanMagic.ConnectSecure;

internal static class ConnectSecurePagedQuery
{
    public const int PageSize = 5000;
    public const int InteractivePageSize = 1000;
    public const int MaxPages = 20;

    /// <summary>
    /// Some report_queries endpoints ignore <c>condition</c> and return tenant-wide pages.
    /// Filter each page by company_id client-side and stop once a page has no rows for the company.
    /// </summary>
    public static async Task<List<JsonElement>> FetchCompanyScopedPagesAsync(
        Func<Dictionary<string, string>, CancellationToken, Task<List<JsonElement>>> fetchPage,
        Dictionary<string, string> baseQuery,
        int companyId,
        CancellationToken ct,
        int pageSize = PageSize,
        int maxPages = MaxPages)
    {
        if (companyId <= 0)
            return [];

        var results = new List<JsonElement>();
        var sawCompanyRows = false;

        for (var page = 0; page < maxPages; page++)
        {
            ct.ThrowIfCancellationRequested();

            var query = new Dictionary<string, string>(baseQuery)
            {
                ["limit"] = pageSize.ToString(),
                ["skip"] = (page * pageSize).ToString()
            };

            var sw = Stopwatch.StartNew();
            var batch = await fetchPage(query, ct);
            sw.Stop();
            if (batch.Count == 0)
                break;

            var companyRows = batch
                .Where(row => ConnectSecureJsonReader.GetInt(row, "company_id", "companyId") == companyId)
                .ToList();

            ConnectSecureRequestMetrics.LogPageFetch(
                "company_scoped",
                page,
                batch.Count,
                companyRows.Count,
                sw.ElapsedMilliseconds);

            if (companyRows.Count > 0)
            {
                sawCompanyRows = true;
                results.AddRange(companyRows);
            }
            else if (sawCompanyRows)
            {
                break;
            }

            if (batch.Count < pageSize)
                break;
        }

        return results;
    }

    /// <summary>
    /// CyberCNS uses skip as page index (0, 1, 2...) on many report_queries endpoints.
    /// Stops when a page repeats (skip ignored) or the batch is empty.
    /// </summary>
    public static async Task<List<JsonElement>> FetchCompanyScopedPagesByIndexAsync(
        Func<Dictionary<string, string>, CancellationToken, Task<List<JsonElement>>> fetchPage,
        Dictionary<string, string> baseQuery,
        int companyId,
        CancellationToken ct,
        int pageSize = InteractivePageSize,
        int maxPages = MaxPages,
        int startPage = 0,
        Action<int>? onFirstCompanyPageFound = null)
    {
        if (companyId <= 0)
            return [];

        var results = new List<JsonElement>();
        var sawCompanyRows = false;
        var firstPageReported = false;
        string? previousFingerprint = null;
        var endPage = startPage + maxPages;

        for (var page = startPage; page < endPage; page++)
        {
            ct.ThrowIfCancellationRequested();

            var query = new Dictionary<string, string>(baseQuery)
            {
                ["limit"] = pageSize.ToString(),
                ["skip"] = page.ToString()
            };

            var sw = Stopwatch.StartNew();
            var batch = await fetchPage(query, ct);
            sw.Stop();
            if (batch.Count == 0)
                break;

            var fingerprint = BatchFingerprint(batch);
            if (fingerprint == previousFingerprint)
                break;
            previousFingerprint = fingerprint;

            var companyRows = batch
                .Where(row => ConnectSecureJsonReader.GetInt(row, "company_id", "companyId") == companyId)
                .ToList();

            ConnectSecureRequestMetrics.LogPageFetch(
                "company_scoped_index",
                page,
                batch.Count,
                companyRows.Count,
                sw.ElapsedMilliseconds);

            if (companyRows.Count > 0)
            {
                if (!firstPageReported)
                {
                    onFirstCompanyPageFound?.Invoke(page);
                    firstPageReported = true;
                }

                sawCompanyRows = true;
                results.AddRange(companyRows);
            }
            else if (sawCompanyRows)
            {
                break;
            }

            if (batch.Count < pageSize)
                break;
        }

        return results;
    }

    public static async Task<List<JsonElement>> FetchAllPagesByIndexAsync(
        Func<Dictionary<string, string>, CancellationToken, Task<List<JsonElement>>> fetchPage,
        Dictionary<string, string> baseQuery,
        CancellationToken ct,
        int pageSize = InteractivePageSize,
        int maxPages = MaxPages)
    {
        var results = new List<JsonElement>();
        string? previousFingerprint = null;

        for (var page = 0; page < maxPages; page++)
        {
            ct.ThrowIfCancellationRequested();

            var query = new Dictionary<string, string>(baseQuery)
            {
                ["limit"] = pageSize.ToString(),
                ["skip"] = page.ToString()
            };

            var sw = Stopwatch.StartNew();
            var batch = await fetchPage(query, ct);
            sw.Stop();
            if (batch.Count == 0)
                break;

            var fingerprint = BatchFingerprint(batch);
            if (fingerprint == previousFingerprint)
                break;
            previousFingerprint = fingerprint;

            ConnectSecureRequestMetrics.LogPageFetch("all_pages_index", page, batch.Count, batch.Count, sw.ElapsedMilliseconds);
            results.AddRange(batch);

            if (batch.Count < pageSize)
                break;
        }

        return results;
    }

    internal static string BatchFingerprint(IReadOnlyList<JsonElement> batch)
    {
        if (batch.Count == 0)
            return "";

        var first = ConnectSecureJsonReader.GetInt(batch[0], "solution_id", "solutionId", "job_id", "jobId", "id") ?? 0;
        var last = ConnectSecureJsonReader.GetInt(batch[^1], "solution_id", "solutionId", "job_id", "jobId", "id") ?? 0;
        return $"{batch.Count}:{first}:{last}";
    }

    public static async Task<List<JsonElement>> FetchAllPagesAsync(
        Func<Dictionary<string, string>, CancellationToken, Task<List<JsonElement>>> fetchPage,
        Dictionary<string, string> baseQuery,
        CancellationToken ct)
    {
        var results = new List<JsonElement>();
        for (var page = 0; page < MaxPages; page++)
        {
            ct.ThrowIfCancellationRequested();

            var query = new Dictionary<string, string>(baseQuery)
            {
                ["limit"] = PageSize.ToString(),
                ["skip"] = (page * PageSize).ToString()
            };

            var sw = Stopwatch.StartNew();
            var batch = await fetchPage(query, ct);
            sw.Stop();
            if (batch.Count == 0)
                break;

            ConnectSecureRequestMetrics.LogPageFetch("all_pages", page, batch.Count, batch.Count, sw.ElapsedMilliseconds);
            results.AddRange(batch);
            if (batch.Count < PageSize)
                break;
        }

        return results;
    }
}
