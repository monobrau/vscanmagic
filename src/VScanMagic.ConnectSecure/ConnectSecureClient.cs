using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using VScanMagic.Core.Models;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureClient
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private static readonly Regex GuidJobIdPattern = new(@"^[a-fA-F0-9-]{36}$", RegexOptions.Compiled);
    private static readonly Regex NumericJobIdPattern = new(@"^\d+$", RegexOptions.Compiled);
    private static readonly HttpClient DownloadHttp = new() { Timeout = TimeSpan.FromMinutes(30) };
    private static readonly SemaphoreSlim DownloadFileLock = new(1, 1);

    private readonly HttpClient _http;
    private readonly RateLimiter _rateLimiter;
    private readonly ConnectSecureOptions _options;
    private readonly ConnectSecureCacheService _cache;

    private readonly SemaphoreSlim _stateLock = new(1, 1);

    private string? _accessToken;
    private string? _userId;
    private DateTimeOffset _tokenExpiry = DateTimeOffset.MinValue;
    private ConnectSecureCredentials? _credentials;

    public ConnectSecureClient(
        HttpClient http,
        RateLimiter rateLimiter,
        ConnectSecureOptions options,
        ConnectSecureCacheService cache)
    {
        _http = http;
        _rateLimiter = rateLimiter;
        _options = options;
        _cache = cache;
    }

    public bool IsConfigured => _credentials is not null &&
        !string.IsNullOrWhiteSpace(_credentials.BaseUrl) &&
        !string.IsNullOrWhiteSpace(_credentials.TenantName) &&
        !string.IsNullOrWhiteSpace(_credentials.ClientId) &&
        !string.IsNullOrWhiteSpace(_credentials.ClientSecret);

    public void Configure(ConnectSecureCredentials credentials)
    {
        var normalized = NormalizeCredentials(credentials);
        if (!_stateLock.Wait(TimeSpan.FromSeconds(30)))
            throw new TimeoutException("Timed out waiting for ConnectSecure client state lock.");

        try
        {
            if (_credentials is not null && CredentialsEqual(_credentials, normalized))
                return;

            _credentials = normalized;
            _accessToken = null;
            _userId = null;
            _tokenExpiry = DateTimeOffset.MinValue;
        }
        finally
        {
            _stateLock.Release();
        }
    }

    public async Task TestAuthenticationAsync(CancellationToken ct = default)
    {
        await _stateLock.WaitAsync(ct);
        try
        {
            _accessToken = null;
            _userId = null;
            _tokenExpiry = DateTimeOffset.MinValue;
        }
        finally
        {
            _stateLock.Release();
        }

        await EnsureAuthenticatedAsync(ct);
    }

    public async Task<IReadOnlyList<CompanyInfo>> GetCompaniesAsync(CancellationToken ct = default)
    {
        if (_cache.TryGet<IReadOnlyList<CompanyInfo>>("companies", out var cached) && cached is not null)
            return cached;

        var response = await InvokeAsync(HttpMethod.Get, "/r/company/companies",
            new Dictionary<string, string> { ["limit"] = "5000", ["skip"] = "0" }, ct: ct);

        var companies = ExtractCompanyArray(response);
        if (companies.Count > 0 && companies[0].ValueKind == JsonValueKind.True)
        {
            var fromStats = await TryGetCompaniesFromStatsAsync(ct);
            if (fromStats.Count > 0)
                return fromStats;

            return Enumerable.Range(1, companies.Count)
                .Select(i => new CompanyInfo(i.ToString(), ""))
                .ToList();
        }

        var parsed = companies
            .Select(ParseCompanyInfo)
            .Where(c => !string.IsNullOrWhiteSpace(c.Id))
            .OrderBy(c => c.Name)
            .ToList();
        _cache.Set("companies", parsed, ConnectSecureCacheService.CompaniesTtl);
        return parsed;
    }

    public Task<IReadOnlyList<StandardReportDescriptor>> GetStandardReportsAsync(
        ReportCatalogScope scope,
        int companyId = 0,
        CancellationToken ct = default) =>
        GetStandardReportsAsync(companyId, scope == ReportCatalogScope.Global, ct);

    public async Task<IReadOnlyList<StandardReportDescriptor>> GetStandardReportsAsync(
        int companyId = 0,
        bool globalOnly = false,
        CancellationToken ct = default)
    {
        var collected = new List<StandardReportDescriptor>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        IEnumerable<bool> globalFlags = globalOnly || companyId == 0
            ? [true]
            : [false, true];

        foreach (var useGlobal in globalFlags)
        {
            foreach (var endpoint in new[] { "/report_builder/standard_reports", "/r/report_builder/standard_reports" })
            {
                try
                {
                    var response = await InvokeAsync(HttpMethod.Get, endpoint,
                        new Dictionary<string, string>
                        {
                            ["isGlobal"] = useGlobal.ToString().ToLowerInvariant(),
                            ["skip"] = "0",
                            ["limit"] = "2000"
                        }, ct: ct);

                    if (IsApiStatusFalse(response))
                        continue;

                    ParseStandardReportsResponse(response, collected, seen);
                    if (collected.Count > 0)
                        return collected;
                }
                catch
                {
                    // try next endpoint / isGlobal
                }
            }
        }

        return collected;
    }

    public string? ResolveStandardReportId(
        string internalType,
        string format,
        IReadOnlyList<StandardReportDescriptor> catalog,
        int companyId = 0)
    {
        var wantFormat = format.ToLowerInvariant();
        if (!StandardReportCatalog.CategoryPatterns.TryGetValue(internalType, out var pattern))
            return null;

        var match = catalog.FirstOrDefault(r =>
            r.ReportType.Equals(wantFormat, StringComparison.OrdinalIgnoreCase) &&
            r.Category.Contains(pattern, StringComparison.OrdinalIgnoreCase));
        if (match is not null) return match.Id;

        match = catalog.FirstOrDefault(r =>
            r.ReportType.Equals(wantFormat, StringComparison.OrdinalIgnoreCase) &&
            r.DisplayName.Contains(pattern, StringComparison.OrdinalIgnoreCase));
        if (match is not null) return match.Id;

        if (StandardReportCatalog.KnownReportIds.TryGetValue(internalType, out var knownId) &&
            catalog.Any(r => r.Id == knownId))
            return knownId;

        var formatMatches = catalog.Where(r => r.ReportType.Equals(wantFormat, StringComparison.OrdinalIgnoreCase)).ToList();
        return formatMatches.Count == 1 ? formatMatches[0].Id : formatMatches.FirstOrDefault()?.Id;
    }

    public async Task<string> CreateReportJobAsync(
        string reportId,
        int companyId,
        string format,
        string reportName,
        string clientName,
        CancellationToken ct = default)
    {
        object companyIdParam = companyId == 0 ? "global" : companyId;
        var companyName = await ResolveCompanyNameAsync(companyId, clientName, ct);
        var reportNameCompact = string.Concat(reportName.Where(c => !char.IsWhiteSpace(c)));
        if (string.IsNullOrWhiteSpace(reportNameCompact)) reportNameCompact = "Report";

        var bodyPortal = new Dictionary<string, object?>
        {
            ["reportId"] = reportId,
            ["reportName"] = reportNameCompact,
            ["reportType"] = "Standard",
            ["isFilter"] = true,
            ["fileType"] = format,
            ["reportFilter"] = new { },
            ["company_id"] = companyIdParam,
            ["company_name"] = companyName
        };

        var bodySnake = new Dictionary<string, object?> { ["company_id"] = companyIdParam, ["report_format"] = format };
        if (Regex.IsMatch(reportId, @"^[a-fA-F0-9]{32}$"))
            bodySnake["report_id"] = reportId;
        else
            bodySnake["reportType"] = reportId;

        var bodyCamel = new Dictionary<string, object?>
        {
            ["company_id"] = companyIdParam,
            ["reportId"] = reportId,
            ["reportType"] = format,
            ["fileType"] = format,
            ["reportName"] = reportName
        };

        string? lastError = null;
        foreach (var endpoint in new[] { "/report_builder/create_report_job", "/r/report_builder/create_report_job" })
        {
            foreach (var body in new[] { bodyPortal, bodySnake, bodyCamel })
            {
                try
                {
                    var response = await InvokeAsync(HttpMethod.Post, endpoint, body: body, ct: ct);
                    if (IsApiStatusFalse(response))
                    {
                        lastError = GetApiErrorMessage(response);
                        if (lastError?.Contains("Please Contact Support", StringComparison.OrdinalIgnoreCase) == true)
                            continue;
                        continue;
                    }

                    var jobId = ExtractJobId(response);
                    if (!string.IsNullOrWhiteSpace(jobId))
                        return jobId;
                }
                catch (Exception ex)
                {
                    lastError = ex.Message;
                    if (lastError.Contains("Please Contact Support", StringComparison.OrdinalIgnoreCase))
                        continue;
                    throw;
                }
            }
        }

        throw new InvalidOperationException(
            "ConnectSecure create_report_job failed. " +
            (lastError ?? "Try Download by Job ID for reports created in the portal."));
    }

    public async Task<string?> GetReportDownloadLinkAsync(string jobId, bool isGlobal, int companyId, CancellationToken ct = default)
    {
        var isGlob = isGlobal.ToString().ToLowerInvariant();
        var jobIdArray = $"[\"{jobId}\"]";
        var variants = new List<Dictionary<string, string>>
        {
            new() { ["job_id"] = jobIdArray, ["isGlobal"] = isGlob },
            new() { ["job_id"] = jobId, ["isGlobal"] = isGlob }
        };

        if (!isGlobal && companyId != 0)
        {
            variants.Add(new Dictionary<string, string>
            {
                ["job_id"] = jobIdArray,
                ["isGlobal"] = isGlob,
                ["company_id"] = companyId.ToString()
            });
            variants.Add(new Dictionary<string, string>
            {
                ["job_id"] = jobId,
                ["isGlobal"] = isGlob,
                ["company_id"] = companyId.ToString()
            });
        }

        foreach (var endpoint in new[] { "/report_builder/get_report_link", "/r/report_builder/get_report_link" })
        {
            foreach (var qp in variants)
            {
                try
                {
                    var response = await InvokeAsync(HttpMethod.Get, endpoint, qp, retryCount: 1, ct: ct);
                    if (IsApiStatusFalse(response))
                        continue;

                    var url = ExtractDownloadUrl(response);
                    if (!string.IsNullOrWhiteSpace(url))
                        return NormalizeUrl(url);
                }
                catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
                {
                    return null;
                }
                catch
                {
                    // try next variant
                }
            }
        }

        return null;
    }

    public Task DownloadFileFromUrlAsync(string downloadUrl, string outputPath, CancellationToken ct = default) =>
        DownloadFileFromUrlAsync(downloadUrl, outputPath, maxAttempts: 5, ct);

    public async Task DownloadFileFromUrlAsync(
        string downloadUrl,
        string outputPath,
        int maxAttempts,
        CancellationToken ct = default)
    {
        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var isPresigned = downloadUrl.Contains("r2.cloudflarestorage", StringComparison.OrdinalIgnoreCase) ||
                          downloadUrl.Contains("X-Amz-Signature", StringComparison.OrdinalIgnoreCase);

        await DownloadFileLock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    using var request = new HttpRequestMessage(HttpMethod.Get, downloadUrl);
                    if (!isPresigned)
                    {
                        var (accessToken, userId) = await GetAuthHeadersAsync(ct).ConfigureAwait(false);
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        if (!string.IsNullOrWhiteSpace(userId))
                            request.Headers.TryAddWithoutValidation("X-USER-ID", userId);
                    }

                    using var response = await DownloadHttp
                        .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct)
                        .ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    var tempPath = outputPath + ".part";
                    if (File.Exists(tempPath))
                    {
                        try { File.Delete(tempPath); } catch { /* best effort */ }
                    }

                    await using (var fs = File.Create(tempPath))
                    {
                        await response.Content.CopyToAsync(fs, ct).ConfigureAwait(false);
                    }

                    if (File.Exists(outputPath))
                    {
                        try { File.Delete(outputPath); } catch { /* replaced below */ }
                    }

                    File.Move(tempPath, outputPath);
                    return;
                }
                catch (IOException ex) when (attempt < maxAttempts && IsTransientFileLock(ex))
                {
                    await Task.Delay(TimeSpan.FromSeconds(Math.Min(attempt * 2, 8)), ct).ConfigureAwait(false);
                }
            }
        }
        finally
        {
            DownloadFileLock.Release();
        }
    }

    private static bool IsTransientFileLock(IOException ex) =>
        ex.Message.Contains("being used by another process", StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("used by another process", StringComparison.OrdinalIgnoreCase);

    public Task<JsonElement> InvokeAuthenticatedAsync(
        HttpMethod method,
        string endpoint,
        IReadOnlyDictionary<string, string>? query = null,
        object? body = null,
        CancellationToken ct = default) =>
        InvokeAsync(method, endpoint, query is null ? null : new Dictionary<string, string>(query), body, ct: ct);

    private async Task<JsonElement> InvokeAsync(
        HttpMethod method,
        string endpoint,
        Dictionary<string, string>? query = null,
        object? body = null,
        int retryCount = 3,
        CancellationToken ct = default)
    {
        for (var attempt = 1; attempt <= retryCount; attempt++)
        {
            var (accessToken, userId) = await GetAuthHeadersAsync(ct);
            await _rateLimiter.WaitAsync(_options.RequestsPerMinute, _options.RequestsPerHour, ct);

            var url = BuildUrl(endpoint, query);
            using var request = new HttpRequestMessage(method, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.TryAddWithoutValidation("Content-Type", "application/json");
            if (!string.IsNullOrWhiteSpace(userId))
                request.Headers.TryAddWithoutValidation("X-USER-ID", userId);

            if (method != HttpMethod.Get && body is not null)
                request.Content = JsonContent.Create(body);

            try
            {
                return await ConnectSecureRequestMetrics.TrackAsync(
                    endpoint,
                    query,
                    async () =>
                    {
                        using var response = await _http.SendAsync(request, ct);
                        var responseBody = await response.Content.ReadAsStringAsync(ct);

                        if (response.StatusCode == HttpStatusCode.TooManyRequests && attempt < retryCount)
                            throw new ConnectSecureRetryException(HttpStatusCode.TooManyRequests);

                        if (response.StatusCode == HttpStatusCode.Unauthorized && attempt < retryCount)
                            throw new ConnectSecureRetryException(HttpStatusCode.Unauthorized);

                        if (response.StatusCode == HttpStatusCode.BadGateway && attempt < retryCount)
                            throw new ConnectSecureRetryException(HttpStatusCode.BadGateway);

                        if (!response.IsSuccessStatusCode)
                            throw new HttpRequestException(
                                $"ConnectSecure request failed ({(int)response.StatusCode}): {Truncate(responseBody, 500)}",
                                null,
                                response.StatusCode);

                        if (string.IsNullOrWhiteSpace(responseBody))
                            return JsonSerializer.SerializeToElement(new { status = true, data = Array.Empty<object>() }, JsonOptions);

                        return JsonSerializer.Deserialize<JsonElement>(responseBody, JsonOptions);
                    });
            }
            catch (ConnectSecureRetryException ex) when (attempt < retryCount)
            {
                if (ex.StatusCode == HttpStatusCode.TooManyRequests)
                {
                    await Task.Delay(TimeSpan.FromSeconds(60), ct);
                    continue;
                }

                if (ex.StatusCode == HttpStatusCode.Unauthorized)
                {
                    await InvalidateAuthAsync(ct);
                    continue;
                }

                if (ex.StatusCode == HttpStatusCode.BadGateway)
                {
                    await Task.Delay(TimeSpan.FromSeconds(5), ct);
                    continue;
                }
            }
        }

        throw new InvalidOperationException($"ConnectSecure request failed after {retryCount} attempts: {endpoint}");
    }

    private sealed class ConnectSecureRetryException(HttpStatusCode statusCode) : Exception
    {
        public HttpStatusCode StatusCode { get; } = statusCode;
    }

    private string BuildUrl(string endpoint, Dictionary<string, string>? query)
    {
        var baseUrl = _credentials!.BaseUrl.TrimEnd('/');
        var path = endpoint.StartsWith("http", StringComparison.OrdinalIgnoreCase)
            ? endpoint
            : baseUrl + endpoint;

        if (query is null || query.Count == 0)
            return path;

        var qs = string.Join("&", query.Select(kv =>
            $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));
        return path + "?" + qs;
    }

    private async Task<(string AccessToken, string? UserId)> GetAuthHeadersAsync(CancellationToken ct)
    {
        await EnsureAuthenticatedAsync(ct);
        await _stateLock.WaitAsync(ct);
        try
        {
            return (_accessToken!, _userId);
        }
        finally
        {
            _stateLock.Release();
        }
    }

    private async Task InvalidateAuthAsync(CancellationToken ct)
    {
        await _stateLock.WaitAsync(ct);
        try
        {
            _accessToken = null;
            _userId = null;
            _tokenExpiry = DateTimeOffset.MinValue;
        }
        finally
        {
            _stateLock.Release();
        }
    }

    private async Task EnsureAuthenticatedAsync(CancellationToken ct)
    {
        await _stateLock.WaitAsync(ct);
        try
        {
            if (_credentials is null)
                throw new InvalidOperationException("ConnectSecure credentials are not configured.");

            if (!string.IsNullOrWhiteSpace(_accessToken) && DateTimeOffset.UtcNow < _tokenExpiry)
                return;

            await AuthenticateCoreAsync(ct);
        }
        finally
        {
            _stateLock.Release();
        }
    }

    private async Task AuthenticateCoreAsync(CancellationToken ct)
    {
        const int maxAuthRetries = 3;
        const int apiErrorRetries = 3;
        string? lastError = null;

        for (var outer = 1; outer <= apiErrorRetries; outer++)
        {
            for (var attempt = 1; attempt <= maxAuthRetries; attempt++)
            {
                await _rateLimiter.WaitAsync(_options.RequestsPerMinute, _options.RequestsPerHour, ct);

                var authString = $"{_credentials!.TenantName}+{_credentials.ClientId}:{_credentials.ClientSecret}";
                var base64Auth = Convert.ToBase64String(Encoding.UTF8.GetBytes(authString));
                var tokenUrl = $"{_credentials.BaseUrl}/w/authorize";

                using var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.TryAddWithoutValidation("Content-Type", "application/json");
                request.Headers.TryAddWithoutValidation("Client-Auth-Token", base64Auth);
                request.Content = new StringContent("{}", Encoding.UTF8, "application/json");

                using var response = await _http.SendAsync(request, ct);
                var content = await response.Content.ReadAsStringAsync(ct);

                if (response.StatusCode == HttpStatusCode.BadGateway && attempt < maxAuthRetries)
                {
                    await Task.Delay(TimeSpan.FromSeconds(5), ct);
                    continue;
                }

                if (response.StatusCode == HttpStatusCode.GatewayTimeout && attempt < maxAuthRetries)
                {
                    await Task.Delay(TimeSpan.FromSeconds(10), ct);
                    continue;
                }

                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(
                        $"ConnectSecure authentication HTTP {(int)response.StatusCode}: {Truncate(content, 500)}");

                var doc = JsonSerializer.Deserialize<JsonElement>(content, JsonOptions);
                if (TryParseAuthResponse(doc, out var accessToken, out var userId, out var errorMessage))
                {
                    _accessToken = accessToken;
                    _userId = userId;
                    _tokenExpiry = DateTimeOffset.UtcNow.AddHours(1);
                    return;
                }

                lastError = errorMessage;
                var messageText = doc.TryGetProperty("message", out var msgEl) && msgEl.ValueKind == JsonValueKind.String
                    ? msgEl.GetString()
                    : null;

                if (messageText == "Failed to create customer" && outer < apiErrorRetries)
                {
                    await Task.Delay(TimeSpan.FromSeconds(8), ct);
                    break;
                }

                if (messageText == "Failed to authorize")
                    throw new InvalidOperationException(lastError ?? "ConnectSecure authentication failed: Failed to authorize.");

                if (outer < apiErrorRetries && messageText == "Failed to create customer")
                    continue;

                throw new InvalidOperationException(lastError ?? "ConnectSecure authentication failed.");
            }
        }

        throw new InvalidOperationException(lastError ?? "ConnectSecure authentication failed.");
    }

    private static ConnectSecureCredentials NormalizeCredentials(ConnectSecureCredentials credentials) =>
        new()
        {
            BaseUrl = credentials.BaseUrl.Trim().TrimEnd('/'),
            TenantName = credentials.TenantName.Trim(),
            ClientId = credentials.ClientId.Trim(),
            ClientSecret = credentials.ClientSecret.Replace("\r", "").Replace("\n", "").Trim()
        };

    private static bool CredentialsEqual(ConnectSecureCredentials left, ConnectSecureCredentials right) =>
        string.Equals(left.BaseUrl, right.BaseUrl, StringComparison.Ordinal) &&
        string.Equals(left.TenantName, right.TenantName, StringComparison.Ordinal) &&
        string.Equals(left.ClientId, right.ClientId, StringComparison.Ordinal) &&
        string.Equals(left.ClientSecret, right.ClientSecret, StringComparison.Ordinal);

    internal static bool TryParseAuthResponse(
        JsonElement doc,
        out string? accessToken,
        out string? userId,
        out string? errorMessage)
    {
        accessToken = null;
        userId = null;
        errorMessage = null;

        if (doc.TryGetProperty("access_token", out var topToken))
            accessToken = topToken.GetString();
        else if (doc.TryGetProperty("token", out var altToken))
            accessToken = altToken.GetString();
        else if (doc.TryGetProperty("data", out var data))
        {
            if (data.TryGetProperty("access_token", out var nestedToken))
                accessToken = nestedToken.GetString();
            else if (data.TryGetProperty("token", out var nestedAlt))
                accessToken = nestedAlt.GetString();
        }

        if (doc.TryGetProperty("user_id", out var topUser))
            userId = topUser.GetString();
        else if (doc.TryGetProperty("data", out var dataUser) && dataUser.TryGetProperty("user_id", out var nestedUser))
            userId = nestedUser.GetString();

        if (!string.IsNullOrWhiteSpace(accessToken))
            return true;

        if (doc.TryGetProperty("message", out var message))
        {
            var msg = message.ValueKind == JsonValueKind.String ? message.GetString() : null;
            errorMessage = msg is not null
                ? ConnectSecureCredentialsHelper.FormatAuthFailureHelp($"ConnectSecure authentication failed: {msg}")
                : ConnectSecureCredentialsHelper.FormatAuthFailureHelp("ConnectSecure authentication failed.");
        }
        else if (doc.TryGetProperty("status", out var status) && status.ValueKind == JsonValueKind.False)
        {
            errorMessage = "ConnectSecure authentication failed.";
        }
        else
        {
            errorMessage = "ConnectSecure authentication failed: no access token in response.";
        }

        return false;
    }

    private async Task<string> ResolveCompanyNameAsync(int companyId, string clientName, CancellationToken ct)
    {
        if (companyId == 0) return "Global";
        if (!string.IsNullOrWhiteSpace(clientName) &&
            !clientName.Equals("Company", StringComparison.OrdinalIgnoreCase) &&
            !clientName.Equals("All Companies", StringComparison.OrdinalIgnoreCase))
            return clientName.Trim();

        try
        {
            var response = await InvokeAsync(HttpMethod.Get, "/r/company/companies",
                new Dictionary<string, string> { ["limit"] = "5000", ["skip"] = "0" }, ct: ct);
            var match = ExtractCompanyArray(response)
                .Select(ParseCompanyInfo)
                .FirstOrDefault(c => c.Id == companyId.ToString());
            if (match is not null && !string.IsNullOrWhiteSpace(match.Name))
                return match.Name;
        }
        catch
        {
            // use fallback below
        }

        return $"Company {companyId}";
    }

    private async Task<List<CompanyInfo>> TryGetCompaniesFromStatsAsync(CancellationToken ct)
    {
        var response = await InvokeAsync(HttpMethod.Get, "/r/company/company_stats",
            new Dictionary<string, string> { ["limit"] = "1000", ["skip"] = "0" }, ct: ct);

        var results = new List<CompanyInfo>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in ExtractCompanyArray(response))
        {
            if (item.ValueKind == JsonValueKind.True) continue;
            var info = ParseCompanyInfo(item);
            if (string.IsNullOrWhiteSpace(info.Id) || !seen.Add(info.Id)) continue;
            results.Add(info);
        }

        return results;
    }

    private static List<JsonElement> ExtractCompanyArray(JsonElement response)
    {
        if (response.ValueKind == JsonValueKind.Array)
            return response.EnumerateArray().ToList();

        if (response.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Array)
            return data.EnumerateArray().ToList();

        if (response.TryGetProperty("hits", out var hits) &&
            hits.TryGetProperty("hits", out var innerHits) &&
            innerHits.ValueKind == JsonValueKind.Array)
        {
            return innerHits.EnumerateArray()
                .Select(h => h.TryGetProperty("_source", out var src) ? src : h)
                .ToList();
        }

        return [];
    }

    private static CompanyInfo ParseCompanyInfo(JsonElement company)
    {
        var name = GetStringProp(company, "name", "company_name", "companyName", "title");
        if (string.IsNullOrWhiteSpace(name) && company.TryGetProperty("_source", out var source))
            name = GetStringProp(source, "name", "company_name", "companyName", "title");

        var id = GetStringProp(company, "id", "company_id", "companyId", "_id");
        if (string.IsNullOrWhiteSpace(id) && company.TryGetProperty("_source", out var sourceId))
            id = GetStringProp(sourceId, "id", "company_id", "companyId", "_id");

        return new CompanyInfo(id, name);
    }

    private static void ParseStandardReportsResponse(JsonElement response, List<StandardReportDescriptor> collected, HashSet<string> seen)
    {
        if (response.TryGetProperty("message", out var message) && message.ValueKind == JsonValueKind.Array)
        {
            foreach (var sec in message.EnumerateArray())
            {
                if (!sec.TryGetProperty("Reports", out var reportsProp) && !sec.TryGetProperty("reports", out reportsProp))
                    continue;

                foreach (var cat in reportsProp.EnumerateArray())
                {
                    var catDisplay = GetStringProp(cat, "description", "Description");
                    var catLower = catDisplay.ToLowerInvariant();
                    if (!cat.TryGetProperty("reports", out var reps) && !cat.TryGetProperty("Reports", out reps))
                        continue;

                    foreach (var rep in reps.EnumerateArray())
                        TryAddReport(rep, catLower, catDisplay, collected, seen);
                }
            }
        }

        if (collected.Count == 0 && response.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Array)
        {
            foreach (var d in data.EnumerateArray())
                TryAddReport(d, d.TryGetProperty("description", out var desc) ? desc.GetString() ?? "" : "", "", collected, seen);
        }
    }

    private static void TryAddReport(JsonElement rep, string category, string categoryDisplay, List<StandardReportDescriptor> collected, HashSet<string> seen)
    {
        var id = GetStringProp(rep, "id", "reportId");
        var rt = GetStringProp(rep, "reportType", "report_type").ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(id) || rt is not ("xlsx" or "docx" or "pdf"))
            return;

        var key = $"{id}-{rt}";
        if (!seen.Add(key)) return;

        collected.Add(new StandardReportDescriptor
        {
            Id = id,
            ReportType = rt,
            Category = category,
            CategoryDisplay = categoryDisplay,
            DisplayName = GetStringProp(rep, "displayReportName", "displayName", "description", "name")
        });
    }

    private static bool IsApiStatusFalse(JsonElement response)
    {
        if (!response.TryGetProperty("status", out var status))
            return false;

        return status.ValueKind == JsonValueKind.False ||
               (status.ValueKind == JsonValueKind.String &&
                status.GetString()?.Equals("false", StringComparison.OrdinalIgnoreCase) == true);
    }

    private static string? GetApiErrorMessage(JsonElement response)
    {
        if (!response.TryGetProperty("message", out var message))
            return "status=false";

        return message.ValueKind == JsonValueKind.String
            ? message.GetString()
            : message.ToString();
    }

    private static string? ExtractJobId(JsonElement response)
    {
        if (response.TryGetProperty("data", out var data))
        {
            if (data.ValueKind == JsonValueKind.String) return data.GetString();
            if (data.ValueKind is JsonValueKind.Number) return data.ToString();
            foreach (var prop in new[] { "job_id", "jobId", "id" })
                if (data.TryGetProperty(prop, out var v)) return v.ToString();
        }

        foreach (var prop in new[] { "job_id", "jobId", "id" })
            if (response.TryGetProperty(prop, out var v)) return v.ToString();

        if (response.TryGetProperty("message", out var msg))
        {
            if (msg.ValueKind == JsonValueKind.String)
            {
                var s = msg.GetString();
                if (!string.IsNullOrWhiteSpace(s) &&
                    (GuidJobIdPattern.IsMatch(s) || NumericJobIdPattern.IsMatch(s)))
                    return s;
            }
            else if (msg.ValueKind == JsonValueKind.Object && msg.TryGetProperty("job_id", out var nested))
                return nested.ToString();
        }

        return null;
    }

    private static string? ExtractDownloadUrl(JsonElement response)
    {
        if (response.TryGetProperty("message", out var msg) && msg.ValueKind == JsonValueKind.String)
        {
            var fromMessage = TryParseDownloadUrlString(msg.GetString());
            if (fromMessage is not null)
                return fromMessage;
        }

        if (response.TryGetProperty("data", out var data))
        {
            if (data.ValueKind == JsonValueKind.String)
            {
                var fromDataString = TryParseDownloadUrlString(data.GetString());
                if (fromDataString is not null)
                    return fromDataString;
            }
            else if (data.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in new[] { "download_url", "url", "link", "file_url", "fileUrl" })
                {
                    if (data.TryGetProperty(prop, out var v))
                    {
                        var fromProp = TryParseDownloadUrlString(
                            v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString());
                        if (fromProp is not null)
                            return fromProp;
                    }
                }
            }
        }

        return null;
    }

    private static string? TryParseDownloadUrlString(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var trimmed = value.Trim();
        if (trimmed.StartsWith("http", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("/", StringComparison.Ordinal) ||
            trimmed.Contains("r2.cloudflarestorage", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Contains("X-Amz-Signature", StringComparison.OrdinalIgnoreCase))
            return trimmed;

        return null;
    }

    private string NormalizeUrl(string url)
    {
        if (url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            return url;
        return $"{_credentials!.BaseUrl.TrimEnd('/')}/{url.TrimStart('/')}";
    }

    private static string GetStringProp(JsonElement el, params string[] names)
    {
        foreach (var n in names)
        {
            if (el.TryGetProperty(n, out var v))
            {
                var s = v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
                if (!string.IsNullOrWhiteSpace(s)) return s!;
            }
        }
        return "";
    }

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max] + "...";
}

public sealed record CompanyInfo(string Id, string Name);
