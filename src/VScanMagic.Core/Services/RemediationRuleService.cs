using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;
using VScanMagic.Core.Risk;

namespace VScanMagic.Core.Services;

public sealed class RemediationRuleService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly string? _configDir;
    private List<RemediationRule>? _cache;

    public RemediationRuleService(string? configDir = null)
    {
        _configDir = configDir;
    }

    public void InvalidateCache() => _cache = null;

    public IReadOnlyList<RemediationRule> LoadRules()
    {
        if (_cache is not null)
            return _cache;

        var path = VScanMagicPaths.RemediationRulesFile(_configDir);
        List<RemediationRule> rules;

        if (!File.Exists(path))
        {
            rules = RemediationRuleDefaults.GetAll();
            SaveRulesInternal(rules, path);
            WriteDefaultsRevision(path);
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

            var changed = false;
            var beforeCount = rules.Count;
            rules = MergeMissingDefaults(rules);
            if (rules.Count > beforeCount)
                changed = true;

            if (ReadDefaultsRevision(path) < RemediationRuleDefaults.Revision)
            {
                changed |= SyncDefaultContent(rules);
                WriteDefaultsRevision(path);
            }

            if (changed)
                SaveRulesInternal(rules, path);
        }

        _cache = rules;
        return _cache;
    }

    public void SaveRules(IEnumerable<RemediationRule> rules)
    {
        _cache = rules.ToList();
        var path = VScanMagicPaths.RemediationRulesFile(_configDir);
        SaveRulesInternal(_cache, path);
        WriteDefaultsRevision(path);
    }

    private static void SaveRulesInternal(List<RemediationRule> rules, string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(path,
            JsonSerializer.Serialize(rules, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string DefaultsRevisionFile(string rulesPath) =>
        Path.ChangeExtension(rulesPath, ".defaults_revision");

    private static int ReadDefaultsRevision(string rulesPath)
    {
        var revisionPath = DefaultsRevisionFile(rulesPath);
        if (!File.Exists(revisionPath))
            return 0;

        var text = File.ReadAllText(revisionPath).Trim();
        return int.TryParse(text, out var revision) ? revision : 0;
    }

    private static void WriteDefaultsRevision(string rulesPath) =>
        File.WriteAllText(DefaultsRevisionFile(rulesPath), RemediationRuleDefaults.Revision.ToString());

    private static List<RemediationRule> MergeMissingDefaults(List<RemediationRule> rules)
    {
        var existing = rules.Select(r => r.Pattern).ToHashSet(StringComparer.Ordinal);
        foreach (var rule in RemediationRuleDefaults.GetAll())
        {
            if (existing.Add(rule.Pattern))
                rules.Add(CloneRule(rule));
        }

        return rules;
    }

    internal static bool SyncDefaultContent(List<RemediationRule> rules)
    {
        var byPattern = rules.ToDictionary(r => r.Pattern, StringComparer.Ordinal);
        var changed = false;

        foreach (var def in RemediationRuleDefaults.GetAll())
        {
            if (!byPattern.TryGetValue(def.Pattern, out var existing))
                continue;

            if (existing.WordText == def.WordText &&
                existing.TicketText == def.TicketText &&
                existing.GuidanceStyle == def.GuidanceStyle &&
                existing.IsDefault == def.IsDefault)
                continue;

            existing.WordText = def.WordText;
            existing.TicketText = def.TicketText;
            existing.GuidanceStyle = def.GuidanceStyle;
            existing.IsDefault = def.IsDefault;
            changed = true;
        }

        return changed;
    }

    private static RemediationRule CloneRule(RemediationRule rule) => new()
    {
        Pattern = rule.Pattern,
        WordText = rule.WordText,
        TicketText = rule.TicketText,
        IsDefault = rule.IsDefault,
        GuidanceStyle = rule.GuidanceStyle
    };

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
        if (TryGetMatchingRule(productName, out var rule))
        {
            guidance = forWord ? rule.WordText : rule.TicketText;
            return true;
        }

        return false;
    }

    public RemediationGuidanceStyle GetGuidanceStyle(string productName)
    {
        if (TryGetMatchingRule(productName, out var rule) &&
            rule.GuidanceStyle != RemediationGuidanceStyle.Standard)
            return rule.GuidanceStyle;

        return AutoUpdateSoftwareHelper.IsAutoUpdating(productName)
            ? RemediationGuidanceStyle.AutoUpdate
            : RemediationGuidanceStyle.Standard;
    }

    public bool PrefersRuleGuidanceOverConnectSecureFix(string productName) =>
        GetGuidanceStyle(productName) == RemediationGuidanceStyle.AutoUpdate;

    public bool TryGetMatchingRule(string productName, out RemediationRule rule)
    {
        rule = null!;
        var rules = LoadRules();
        foreach (var candidate in rules
                     .Where(r => !r.IsDefault && r.Pattern != "*")
                     .OrderByDescending(r => r.Pattern.Length))
        {
            if (!MatchesWildcard(productName, candidate.Pattern))
                continue;

            rule = candidate;
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
