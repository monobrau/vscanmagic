namespace VScanMagic.ConnectWiseManage;

public sealed class ConnectWiseManageCredentials
{
    public string ApiUrl { get; set; } = "";
    /// <summary>Manage member company identifier used in API auth (companyId+publicKey).</summary>
    public string CompanyId { get; set; } = "";
    public string PublicKey { get; set; } = "";
    public string PrivateKey { get; set; } = "";
    public string ClientId { get; set; } = "";
}

public sealed class ConnectWiseManageOptions
{
    public int DefaultBoardId { get; set; }
    public int DefaultStatusId { get; set; }
}

public sealed class ConnectWiseCompanyMapEntry
{
    public int ManageCompanyId { get; set; }
    public string ManageCompanyName { get; set; } = "";
}
