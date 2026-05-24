using System.Text.Json;

namespace VScanMagic.ConnectSecure;

internal static class ConnectSecurePagedQuery
{
    public const int PageSize = 5000;
    public const int MaxPages = 20;

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

            var batch = await fetchPage(query, ct);
            if (batch.Count == 0)
                break;

            results.AddRange(batch);
            if (batch.Count < PageSize)
                break;
        }

        return results;
    }
}
