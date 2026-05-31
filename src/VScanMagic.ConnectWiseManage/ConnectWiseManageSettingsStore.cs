using System.Text.Json;
using VScanMagic.Core.Paths;

namespace VScanMagic.ConnectWiseManage;

public sealed class ConnectWiseManageSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, PropertyNameCaseInsensitive = true };

    public ConnectWiseManageCredentials LoadCredentials()
    {
        var path = VScanMagicPaths.ConnectWiseManageCredentialsFile();
        if (!File.Exists(path))
            return new ConnectWiseManageCredentials();

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<ConnectWiseManageCredentials>(json, JsonOptions)
                ?? new ConnectWiseManageCredentials();
        }
        catch
        {
            return new ConnectWiseManageCredentials();
        }
    }

    public void SaveCredentials(ConnectWiseManageCredentials credentials)
    {
        var dir = VScanMagicPaths.GetConfigDirectory();
        Directory.CreateDirectory(dir);
        File.WriteAllText(
            VScanMagicPaths.ConnectWiseManageCredentialsFile(),
            JsonSerializer.Serialize(credentials, JsonOptions));
    }

    public ConnectWiseManageOptions LoadOptions()
    {
        var path = VScanMagicPaths.ConnectWiseManageOptionsFile();
        if (!File.Exists(path))
            return new ConnectWiseManageOptions();

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<ConnectWiseManageOptions>(json, JsonOptions)
                ?? new ConnectWiseManageOptions();
        }
        catch
        {
            return new ConnectWiseManageOptions();
        }
    }

    public void SaveOptions(ConnectWiseManageOptions options)
    {
        var dir = VScanMagicPaths.GetConfigDirectory();
        Directory.CreateDirectory(dir);
        File.WriteAllText(
            VScanMagicPaths.ConnectWiseManageOptionsFile(),
            JsonSerializer.Serialize(options, JsonOptions));
    }
}
