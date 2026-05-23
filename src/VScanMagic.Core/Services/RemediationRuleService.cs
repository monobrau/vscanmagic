using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class RemediationRuleService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private List<RemediationRule>? _cache;

    public IReadOnlyList<RemediationRule> LoadRules()
    {
        if (_cache is not null)
            return _cache;

        var path = VScanMagicPaths.RemediationRulesFile();
        List<RemediationRule> rules;

        if (!File.Exists(path))
        {
            rules = RemediationRuleDefaults.GetAll();
            SaveRulesInternal(rules, path);
        }
        else
        {
            try
            {
                var json = File.ReadAllText(path);
                rules = JsonSerializer.Deserialize<List<RemediationRule>>(json, JsonOptions) ?? [];
            }
            catch
            {
                rules = [];
            }

            var beforeCount = rules.Count;
            rules = MergeMissingDefaults(rules);
            if (rules.Count > beforeCount)
                SaveRulesInternal(rules, path);
        }

        _cache = rules;
        return _cache;
    }

    public void SaveRules(IEnumerable<RemediationRule> rules)
    {
        _cache = rules.ToList();
        SaveRulesInternal(_cache, VScanMagicPaths.RemediationRulesFile());
    }

    private static void SaveRulesInternal(List<RemediationRule> rules, string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path,
            JsonSerializer.Serialize(rules, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static List<RemediationRule> MergeMissingDefaults(List<RemediationRule> rules)
    {
        var existing = rules.Select(r => r.Pattern).ToHashSet(StringComparer.Ordinal);
        foreach (var rule in RemediationRuleDefaults.GetAll())
        {
            if (existing.Add(rule.Pattern))
                rules.Add(rule);
        }

        return rules;
    }

    public string GetGuidance(string productName, bool forWord = true)
    {
        if (TryGetSpecificGuidance(productName, forWord, out var guidance))
            return guidance;

        var rules = LoadRules();
        var defaultRule = rules.FirstOrDefault(r => r.IsDefault || r.Pattern == "*");
        if (defaultRule is not null)
            return forWord ? defaultRule.WordText : defaultRule.TicketText;

        return forWord
            ? "This application should be updated to the latest version. If available via ConnectWise Automate/RMM or scripting, deploy updates using the patch management system or scripts. Otherwise, manual updates may be required on affected systems."
            : "- Update to latest version\r\n  - Deploy via ConnectWise Automate/RMM or scripting if available\r\n  - Otherwise, manual updates required on affected systems";
    }

    public bool TryGetSpecificGuidance(string productName, bool forWord, out string guidance)
    {
        guidance = "";
        var rules = LoadRules();
        foreach (var rule in rules
                     .Where(r => !r.IsDefault && r.Pattern != "*")
                     .OrderByDescending(r => r.Pattern.Length))
        {
            if (!MatchesWildcard(productName, rule.Pattern))
                continue;

            guidance = forWord ? rule.WordText : rule.TicketText;
            return true;
        }

        return false;
    }

    private static bool MatchesWildcard(string input, string pattern)
    {
        if (string.IsNullOrEmpty(pattern) || pattern == "*")
            return true;

        if (pattern.Contains('*'))
        {
            var parts = pattern.Split('*', StringSplitOptions.RemoveEmptyEntries);
            var idx = 0;
            foreach (var part in parts)
            {
                var found = input.IndexOf(part, idx, StringComparison.OrdinalIgnoreCase);
                if (found < 0) return false;
                idx = found + part.Length;
            }
            return true;
        }

        return input.Contains(pattern, StringComparison.OrdinalIgnoreCase);
    }
}
