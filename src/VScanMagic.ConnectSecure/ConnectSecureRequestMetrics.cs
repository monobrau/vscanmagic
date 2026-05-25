using System.Diagnostics;
using System.Text.Json;
using Serilog;

namespace VScanMagic.ConnectSecure;

internal static class ConnectSecureRequestMetrics
{
    public static async Task<JsonElement> TrackAsync(
        string endpoint,
        IReadOnlyDictionary<string, string>? query,
        Func<Task<JsonElement>> action)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            var result = await action();
            sw.Stop();
            var rowCount = TryCountRows(result);
            Log.Information(
                "ConnectSecure {Endpoint} {ElapsedMs}ms rows={RowCount} query={Query}",
                endpoint,
                sw.ElapsedMilliseconds,
                rowCount,
                FormatQuery(query));
            return result;
        }
        catch (Exception ex)
        {
            sw.Stop();
            Log.Warning(
                ex,
                "ConnectSecure {Endpoint} failed after {ElapsedMs}ms query={Query}",
                endpoint,
                sw.ElapsedMilliseconds,
                FormatQuery(query));
            throw;
        }
    }

    public static void LogPageFetch(string context, int page, int batchRows, int keptRows, long elapsedMs) =>
        Log.Debug(
            "ConnectSecure page {Context} page={Page} batch={BatchRows} kept={KeptRows} {ElapsedMs}ms",
            context,
            page,
            batchRows,
            keptRows,
            elapsedMs);

    private static int TryCountRows(JsonElement response)
    {
        if (response.ValueKind == JsonValueKind.Array)
            return response.GetArrayLength();

        if (response.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Array)
            return data.GetArrayLength();

        return -1;
    }

    private static string FormatQuery(IReadOnlyDictionary<string, string>? query)
    {
        if (query is null || query.Count == 0)
            return "";

        return string.Join("&", query.Select(kv => $"{kv.Key}={kv.Value}"));
    }
}
