namespace VScanMagic.ConnectSecure;

public static class CredentialTypeCatalog
{
    private static readonly CredentialTypeDefinition Generic = new("generic", "Generic", []);

    private static readonly IReadOnlyList<CredentialTypeDefinition> Defaults =
    [
        new("windows", "Windows / Domain", [
            new("username", "Username", false, "DOMAIN\\user"),
            new("password", "Password", true, null),
            new("domain", "Domain", false, "CONTOSO"),
            new("port", "Port", false, "445")
        ]),
        new("linux", "Linux / SSH", [
            new("username", "Username", false, "root"),
            new("password", "Password", true, null),
            new("port", "Port", false, "22"),
            new("private_key", "Private key", true, null)
        ]),
        new("ssh", "SSH", [
            new("username", "Username", false, "admin"),
            new("password", "Password", true, null),
            new("port", "Port", false, "22"),
            new("private_key", "Private key", true, null)
        ]),
        new("snmp", "SNMP", [
            new("community", "Community string", true, "public"),
            new("version", "Version", false, "2c"),
            new("port", "Port", false, "161")
        ]),
        new("vmware", "VMware", [
            new("username", "Username", false, "administrator@vsphere.local"),
            new("password", "Password", true, null),
            new("host", "Host / vCenter", false, "vcenter.example.com"),
            new("port", "Port", false, "443")
        ]),
        new("mac", "macOS", [
            new("username", "Username", false, "admin"),
            new("password", "Password", true, null)
        ]),
        new("azure", "Azure", [
            new("tenant_id", "Tenant ID", false, null),
            new("client_id", "Client ID", false, null),
            new("client_secret", "Client secret", true, null)
        ]),
        new("aws", "AWS", [
            new("access_key", "Access key", false, null),
            new("secret_key", "Secret key", true, null),
            new("region", "Region", false, "us-east-1")
        ]),
        Generic
    ];

    public static IReadOnlyList<CredentialTypeDefinition> AllDefaults => Defaults;

    public static CredentialTypeDefinition Resolve(string? credentialType)
    {
        if (string.IsNullOrWhiteSpace(credentialType))
            return Generic;

        var match = Defaults.FirstOrDefault(d =>
            string.Equals(d.Type, credentialType, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
            return match;

        return new CredentialTypeDefinition(credentialType.Trim(), credentialType.Trim(), []);
    }

    public static IReadOnlyList<string> MergeKnownTypes(IEnumerable<string> discovered)
    {
        var types = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var def in Defaults)
            types.Add(def.Type);
        foreach (var type in discovered)
        {
            if (!string.IsNullOrWhiteSpace(type))
                types.Add(type.Trim());
        }

        return types.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
    }
}

public sealed record CredentialTypeDefinition(
    string Type,
    string Label,
    IReadOnlyList<CredentialParamField> Fields);

public sealed record CredentialParamField(
    string Key,
    string Label,
    bool IsSecret,
    string? Placeholder);
