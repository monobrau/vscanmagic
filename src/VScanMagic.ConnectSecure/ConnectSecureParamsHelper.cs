using System.Text.Json;
using System.Text.Json.Nodes;

namespace VScanMagic.ConnectSecure;

public static class ConnectSecureParamsHelper
{
    private static readonly HashSet<string> SecretKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "password", "secret", "client_secret", "private_key", "community", "token", "api_key", "access_key", "secret_key"
    };

    public static string SummarizeParams(JsonElement paramsElement)
    {
        if (paramsElement.ValueKind != JsonValueKind.Object)
            return "";

        var parts = new List<string>();
        foreach (var property in paramsElement.EnumerateObject())
        {
            if (string.IsNullOrWhiteSpace(property.Name))
                continue;

            if (IsSecretKey(property.Name))
            {
                parts.Add($"{property.Name}=••••");
                continue;
            }

            var value = property.Value.ValueKind == JsonValueKind.String
                ? property.Value.GetString()
                : property.Value.ToString();
            if (!string.IsNullOrWhiteSpace(value))
                parts.Add($"{property.Name}={value}");
        }

        return string.Join(", ", parts);
    }

    public static JsonObject ParseParamsObject(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return new JsonObject();

        try
        {
            return JsonNode.Parse(json)?.AsObject() ?? new JsonObject();
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException($"Params JSON is invalid: {ex.Message}");
        }
    }

    public static JsonObject MergeParamsFromForm(
        JsonObject existing,
        IReadOnlyDictionary<string, string?> formValues,
        bool mergeExistingSecrets)
    {
        var merged = existing.DeepClone()?.AsObject() ?? new JsonObject();
        foreach (var (key, value) in formValues)
        {
            if (string.IsNullOrWhiteSpace(key))
                continue;

            if (string.IsNullOrWhiteSpace(value))
            {
                if (!mergeExistingSecrets || !IsSecretKey(key))
                    merged.Remove(key);
                continue;
            }

            merged[key] = value;
        }

        return merged;
    }

    public static JsonObject MergeParamsJson(JsonObject existing, JsonObject incoming, bool mergeExistingSecrets)
    {
        var merged = existing.DeepClone()?.AsObject() ?? new JsonObject();
        foreach (var property in incoming)
        {
            if (string.IsNullOrWhiteSpace(property.Key))
                continue;

            if (property.Value is null)
            {
                merged.Remove(property.Key);
                continue;
            }

            var text = property.Value.ToJsonString().Trim('"');
            if (string.IsNullOrWhiteSpace(text) && mergeExistingSecrets && IsSecretKey(property.Key) &&
                merged.TryGetPropertyValue(property.Key, out _))
                continue;

            merged[property.Key] = property.Value.DeepClone();
        }

        return merged;
    }

    public static Dictionary<string, string?> ExtractFormValues(JsonElement paramsElement, CredentialTypeDefinition definition)
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in definition.Fields)
            values[field.Key] = "";

        if (paramsElement.ValueKind != JsonValueKind.Object)
            return values;

        foreach (var field in definition.Fields)
        {
            if (!paramsElement.TryGetProperty(field.Key, out var value))
                continue;

            values[field.Key] = value.ValueKind == JsonValueKind.String
                ? value.GetString()
                : value.ToString();
        }

        return values;
    }

    public static bool IsSecretKey(string key) =>
        SecretKeys.Any(secret => key.Contains(secret, StringComparison.OrdinalIgnoreCase));
}
