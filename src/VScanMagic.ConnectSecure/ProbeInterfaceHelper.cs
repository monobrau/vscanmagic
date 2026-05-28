using System.Text.Json;

namespace VScanMagic.ConnectSecure;

public static class ProbeInterfaceHelper
{
    public static IReadOnlyList<string> ParseAvailableInterfaces(JsonElement agent)
    {
        if (!TryGetInterfacesArray(agent, out var array))
            return [];

        var results = new List<string>();
        foreach (var item in array.EnumerateArray())
        {
            var parsed = ParseInterfaceItem(item);
            if (!string.IsNullOrWhiteSpace(parsed))
                results.Add(parsed);
        }

        return results
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static IReadOnlyList<ProbeInterfaceOption> BuildDropdownOptions(
        IReadOnlyList<string> availableInterfaces,
        string? currentValue)
    {
        var options = new List<ProbeInterfaceOption>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var iface in availableInterfaces)
        {
            if (seen.Add(iface))
                options.Add(new ProbeInterfaceOption(iface, FormatLabel(iface)));
        }

        if (!string.IsNullOrWhiteSpace(currentValue) &&
            !currentValue.Equals("(not set)", StringComparison.OrdinalIgnoreCase) &&
            seen.Add(currentValue))
            options.Add(new ProbeInterfaceOption(currentValue, FormatLabel(currentValue)));

        if (options.Count == 0)
            options.Add(new ProbeInterfaceOption("", "(not set)"));

        return options;
    }

    public static string FormatLabel(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "(not set)";

        return value.Trim();
    }

    private static bool TryGetInterfacesArray(JsonElement agent, out JsonElement array)
    {
        foreach (var name in new[] { "interfaces", "Interfaces", "nmap_interfaces", "nmapInterfaces" })
        {
            if (agent.TryGetProperty(name, out array) && array.ValueKind == JsonValueKind.Array)
                return true;
        }

        array = default;
        return false;
    }

    private static string ParseInterfaceItem(JsonElement item)
    {
        if (item.ValueKind == JsonValueKind.String)
            return item.GetString()?.Trim() ?? "";

        if (item.ValueKind != JsonValueKind.Object)
            return item.ToString().Trim();

        var name = ConnectSecureJsonReader.GetString(item,
            "name", "interface", "device", "adapter", "ifname", "if_name");
        var ip = ConnectSecureJsonReader.GetString(item,
            "ip", "address", "ipv4", "ip_address", "ipAddress");

        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(ip))
            return $"{name} ({ip})";
        if (!string.IsNullOrWhiteSpace(ip))
            return ip;
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        return item.GetRawText();
    }
}

public sealed record ProbeInterfaceOption(string Value, string Label);
