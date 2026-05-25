namespace VScanMagic.Data.Parsing;

public static class ColumnMapper
{
    private static readonly Dictionary<string, string[]> Mappings = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Source"] = ["Source", "Section", "Vulnerability Source"],
        ["HostName"] = ["Asset Name", "Host Name", "Hostname", "Computer", "Computer Name", "Device", "Device Name", "System", "System Name", "Machine", "Asset", "Host", "Endpoint", "Target"],
        ["IP"] = ["IP Address", "IP", "Address"],
        ["Product"] = ["Product Name", "Application Name", "Software Name", "Product", "App Name", "OS Name", "OS Full Name", "Name", "Problem Name"],
        ["Severity"] = ["Severity"],
        ["EPSS"] = ["EPSS Score", "EPSS", "Exploit Prediction Score"],
        ["Fix"] = ["Solution", "Fix", "Remediation", "FIX"],
        ["Username"] = ["Username", "User Name", "User", "Account", "Login", "Login Name", "Last User", "Last Logged In User", "Last Logged In User Name", "Logged In User", "Owner", "Asset Owner", "Primary User"],
        ["CVE"] = ["CVE ID", "CVE", "Problem Name", "problem_name"],
        ["CVSS"] = ["CVSS Score", "CVSS", "Base Score", "base_score"],
        ["AffectedAssets"] = ["Affected Assets", "Affected Assets Count", "affected_assets"],
        ["Critical"] = ["Critical"],
        ["High"] = ["High"],
        ["Medium"] = ["Medium"],
        ["Low"] = ["Low"],
        ["VulnCount"] = ["Vulnerability Count", "Vuln Count", "Count"]
    };

    public static Dictionary<string, int> MapColumns(IReadOnlyList<string> headers)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headerMap = headers
            .Select((h, i) => (Header: h?.Trim() ?? "", Index: i + 1))
            .Where(x => !string.IsNullOrEmpty(x.Header))
            .ToDictionary(x => x.Header, x => x.Index, StringComparer.OrdinalIgnoreCase);

        foreach (var (key, possible) in Mappings)
        {
            var idx = FindColumnIndex(headerMap, possible);
            if (idx.HasValue)
                result[key] = idx.Value;
        }

        return result;
    }

    private static int? FindColumnIndex(Dictionary<string, int> headers, string[] possibleNames)
    {
        foreach (var name in possibleNames)
        {
            if (headers.TryGetValue(name, out var exact))
                return exact;
        }

        foreach (var name in possibleNames)
        {
            foreach (var (header, idx) in headers)
            {
                if (header.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return idx;
            }
        }

        foreach (var name in possibleNames)
        {
            foreach (var (header, idx) in headers)
            {
                if (header.Contains(name, StringComparison.OrdinalIgnoreCase) ||
                    name.Contains(header, StringComparison.OrdinalIgnoreCase))
                    return idx;
            }
        }

        return null;
    }

    public static int GetSafeInt(string? value, int defaultValue = 0)
    {
        if (string.IsNullOrWhiteSpace(value)) return defaultValue;
        var clean = value.Replace(",", "").Replace(" ", "");
        return int.TryParse(clean, out var i) ? i :
            double.TryParse(clean, out var d) ? (int)Math.Round(d) : defaultValue;
    }

    public static double GetSafeDouble(string? value, double defaultValue = 0)
    {
        if (string.IsNullOrWhiteSpace(value)) return defaultValue;
        var clean = value.Replace(",", "").Replace(" ", "");
        return double.TryParse(clean, out var d) ? d : defaultValue;
    }
}
