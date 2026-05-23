using System.Text.Json;

namespace VScanMagic.ConnectSecure;

internal static class ConnectSecureJsonReader
{
    public static string GetString(JsonElement el, params string[] names)
    {
        foreach (var name in names)
        {
            if (!el.TryGetProperty(name, out var value))
                continue;

            var text = value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
            if (!string.IsNullOrWhiteSpace(text))
                return text!;
        }

        return "";
    }

    public static int? GetInt(JsonElement el, params string[] names)
    {
        foreach (var name in names)
        {
            if (!el.TryGetProperty(name, out var value))
                continue;

            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var n))
                return n;
            if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out n))
                return n;
        }

        return null;
    }

    public static bool GetBool(JsonElement el, params string[] names)
    {
        foreach (var name in names)
        {
            if (!el.TryGetProperty(name, out var value))
                continue;

            if (value.ValueKind == JsonValueKind.True) return true;
            if (value.ValueKind == JsonValueKind.False) return false;
            if (value.ValueKind == JsonValueKind.String &&
                bool.TryParse(value.GetString(), out var b))
                return b;
        }

        return false;
    }

    public static List<JsonElement> ExtractDataArray(JsonElement response)
    {
        if (response.ValueKind == JsonValueKind.Array)
            return response.EnumerateArray().ToList();

        if (response.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Array)
            return data.EnumerateArray().ToList();

        if (response.TryGetProperty("message", out var message) && message.ValueKind == JsonValueKind.Array)
            return message.EnumerateArray().ToList();

        return [];
    }
}
