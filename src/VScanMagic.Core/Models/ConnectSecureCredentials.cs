namespace VScanMagic.Core.Models;

public sealed class ConnectSecureCredentials
{
    public string BaseUrl { get; set; } = "";
    public string TenantName { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
}
