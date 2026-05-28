using System.Text;
using System.Text.RegularExpressions;
using VScanMagic.Core.Models;

namespace VScanMagic.ConnectSecure;

public static class ConnectSecureCredentialsHelper
{
    private static readonly Regex Base64Pattern = new(@"^[A-Za-z0-9+/=]+$", RegexOptions.Compiled);

    public static IReadOnlyList<string> Validate(ConnectSecureCredentials credentials)
    {
        var issues = new List<string>();

        if (string.IsNullOrWhiteSpace(credentials.BaseUrl))
            issues.Add("Base URL is required (include https://, e.g. https://pod0.myconnectsecure.com).");
        else if (!credentials.BaseUrl.Trim().StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            issues.Add("Base URL should start with https://.");

        if (string.IsNullOrWhiteSpace(credentials.TenantName))
            issues.Add("Tenant name is required (exact match from ConnectSecure API Key page).");

        if (string.IsNullOrWhiteSpace(credentials.ClientId))
            issues.Add("Client ID is required.");

        if (string.IsNullOrWhiteSpace(credentials.ClientSecret))
            issues.Add("Client Secret is required.");

        return issues;
    }

    /// <summary>
    /// ConnectSecure stores API secrets with Fernet encryption at rest. The value from the API Key page
    /// is typically a long base64 string (often starting with Z0FBQUFB…) and is used as-is in auth.
    /// </summary>
    public static bool IsConnectSecureEncodedSecret(string secret)
    {
        secret = secret.Trim();
        if (secret.Length < 80)
            return false;

        if (secret.StartsWith("Z0FBQUFB", StringComparison.Ordinal) ||
            secret.StartsWith("gAAAAA", StringComparison.Ordinal))
            return true;

        return secret.Length >= 120 && Base64Pattern.IsMatch(secret);
    }

    public static string BuildClientAuthToken(ConnectSecureCredentials credentials)
    {
        var tenant = credentials.TenantName.Trim();
        var clientId = credentials.ClientId.Trim();
        var secret = credentials.ClientSecret.Replace("\r", "").Replace("\n", "").Trim();
        return Convert.ToBase64String(Encoding.UTF8.GetBytes($"{tenant}+{clientId}:{secret}"));
    }

    public static string GetSwaggerUrl(ConnectSecureCredentials credentials) =>
        $"{credentials.BaseUrl.Trim().TrimEnd('/')}/apidocs/";

    public static string FormatAuthFailureHelp(string? apiMessage, ConnectSecureCredentials? credentials = null)
    {
        var lines = new List<string> { apiMessage ?? "ConnectSecure authentication failed." };

        if (apiMessage?.Contains("Failed to authorize", StringComparison.OrdinalIgnoreCase) == true)
        {
            lines.Add("ConnectSecure rejected the credential triple (tenant + client ID + client secret).");
            lines.Add("Use the secret exactly as shown on the API Key page — long base64 values starting with Z0FB… are normal.");
            lines.Add("Re-copy Base URL, Tenant, Client ID, and Client Secret together from the same API Key screen.");
            lines.Add("If the key was regenerated or Reset User Secret was used, create a new API key.");
            lines.Add("Confirm in Swagger (Profile → API Documentation → POST /w/authorize) before retrying here.");
        }
        else
        {
            lines.Add("Verify on ConnectSecure: Global → Settings → Users → API Key.");
            lines.Add("Copy Base URL, Tenant Name, Client ID, and Client Secret exactly (no extra spaces).");
        }

        return string.Join(" ", lines);
    }
}
