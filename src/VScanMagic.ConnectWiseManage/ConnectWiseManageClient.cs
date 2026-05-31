using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VScanMagic.ConnectWiseManage;

public sealed class ConnectWiseManageClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly HttpClient _http;
    private ConnectWiseManageCredentials? _credentials;

    public ConnectWiseManageClient(HttpClient http) => _http = http;

    public bool IsConfigured => _credentials is not null &&
        !string.IsNullOrWhiteSpace(_credentials.ApiUrl) &&
        !string.IsNullOrWhiteSpace(_credentials.CompanyId) &&
        !string.IsNullOrWhiteSpace(_credentials.PublicKey) &&
        !string.IsNullOrWhiteSpace(_credentials.PrivateKey) &&
        !string.IsNullOrWhiteSpace(_credentials.ClientId);

    public void Configure(ConnectWiseManageCredentials credentials)
    {
        _credentials = credentials;
        var baseUrl = credentials.ApiUrl.Trim().TrimEnd('/');
        if (!baseUrl.EndsWith("/apis/3.0", StringComparison.OrdinalIgnoreCase))
            baseUrl += "/apis/3.0";
        _http.BaseAddress = new Uri(baseUrl + "/");

        var authValue = Convert.ToBase64String(
            Encoding.UTF8.GetBytes($"{credentials.CompanyId}+{credentials.PublicKey}:{credentials.PrivateKey}"));
        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authValue);
        _http.DefaultRequestHeaders.Remove("clientId");
        _http.DefaultRequestHeaders.Add("clientId", credentials.ClientId.Trim());
    }

    public async Task<string> TestConnectionAsync(CancellationToken ct = default)
    {
        var boards = await GetBoardsAsync(ct).ConfigureAwait(false);
        return boards.Count == 0
            ? "Connected, but no service boards were returned."
            : $"Connected. {boards.Count} service board(s) available.";
    }

    public async Task<IReadOnlyList<ManageBoard>> GetBoardsAsync(CancellationToken ct = default)
    {
        EnsureConfigured();
        var response = await _http.GetAsync("service/boards?pageSize=250", ct).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        var boards = await response.Content.ReadFromJsonAsync<List<ManageBoard>>(JsonOptions, ct).ConfigureAwait(false);
        return boards ?? [];
    }

    public async Task<IReadOnlyList<ManageStatus>> GetBoardStatusesAsync(int boardId, CancellationToken ct = default)
    {
        EnsureConfigured();
        var response = await _http.GetAsync($"service/boards/{boardId}/statuses?pageSize=250", ct).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        var statuses = await response.Content.ReadFromJsonAsync<List<ManageStatus>>(JsonOptions, ct).ConfigureAwait(false);
        return statuses ?? [];
    }

    public async Task<ManageTicket> CreateTicketAsync(ManageTicketCreateRequest request, CancellationToken ct = default)
    {
        EnsureConfigured();
        using var response = await _http.PostAsJsonAsync("service/tickets", request, JsonOptions, ct).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        var ticket = await response.Content.ReadFromJsonAsync<ManageTicket>(JsonOptions, ct).ConfigureAwait(false);
        return ticket ?? throw new InvalidOperationException("Manage API returned an empty ticket payload.");
    }

    public async Task<ManageTicket> GetTicketAsync(int ticketId, CancellationToken ct = default)
    {
        EnsureConfigured();
        var response = await _http.GetAsync($"service/tickets/{ticketId}", ct).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        var ticket = await response.Content.ReadFromJsonAsync<ManageTicket>(JsonOptions, ct).ConfigureAwait(false);
        return ticket ?? throw new InvalidOperationException($"Ticket {ticketId} was not found.");
    }

    private void EnsureConfigured()
    {
        if (!IsConfigured)
            throw new InvalidOperationException("ConnectWise Manage is not configured.");
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response)
    {
        if (response.IsSuccessStatusCode)
            return;

        var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        throw new InvalidOperationException(
            $"ConnectWise Manage API {(int)response.StatusCode}: {TrimError(body)}");
    }

    private static string TrimError(string body) =>
        body.Length > 500 ? body[..500] + "..." : body;
}

public sealed class ManageBoard
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
}

public sealed class ManageStatus
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
}

public sealed class ManageTicket
{
    public int Id { get; set; }
    public string? Summary { get; set; }
    public ManageStatusReference? Status { get; set; }

    /// <summary>Service board ticket number when returned separately from the record id.</summary>
    [JsonPropertyName("ticketNumber")]
    public int? BoardTicketNumber { get; set; }

    public string GetDisplayTicketNumber()
    {
        if (BoardTicketNumber is > 0)
            return BoardTicketNumber.Value.ToString();

        return Id > 0 ? Id.ToString() : "";
    }
}

public sealed class ManageStatusReference
{
    public int Id { get; set; }
    public string? Name { get; set; }
}

public sealed class ManageReference
{
    public int Id { get; set; }
}

public sealed class ManageTicketCreateRequest
{
    public string Summary { get; set; } = "";
    public string InitialDescription { get; set; } = "";
    public ManageReference Board { get; set; } = new();
    public ManageReference Status { get; set; } = new();
    public ManageReference Company { get; set; } = new();
}
