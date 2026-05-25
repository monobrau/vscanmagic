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
