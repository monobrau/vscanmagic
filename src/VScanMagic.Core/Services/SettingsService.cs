using System.Text.Json;
using VScanMagic.Core.Models;
using VScanMagic.Core.Paths;

namespace VScanMagic.Core.Services;

public sealed class SettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, PropertyNameCaseInsensitive = true };

    public UserSettings LoadUserSettings()
    {
        var path = VScanMagicPaths.SettingsFile();
        if (!File.Exists(path))
            return new UserSettings();

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<UserSettings>(json, JsonOptions) ?? new UserSettings();
        }
        catch
        {
            return new UserSettings();
        }
    }

    public void SaveUserSettings(UserSettings settings)
    {
        var dir = VScanMagicPaths.GetConfigDirectory(settings.SettingsDirectory);
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "VScanMagic_Settings.json");
        File.WriteAllText(path, JsonSerializer.Serialize(settings, JsonOptions));
    }

    public ConnectSecureCredentials LoadConnectSecureCredentials()
    {
        var path = VScanMagicPaths.ConnectSecureCredentialsFile();
        if (!File.Exists(path))
            return new ConnectSecureCredentials();

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<ConnectSecureCredentials>(json, JsonOptions) ?? new ConnectSecureCredentials();
        }
        catch
        {
            return new ConnectSecureCredentials();
        }
    }

    public void SaveConnectSecureCredentials(ConnectSecureCredentials credentials)
    {
        var dir = VScanMagicPaths.GetConfigDirectory();
        Directory.CreateDirectory(dir);
        File.WriteAllText(VScanMagicPaths.ConnectSecureCredentialsFile(), JsonSerializer.Serialize(credentials, JsonOptions));
    }
}
