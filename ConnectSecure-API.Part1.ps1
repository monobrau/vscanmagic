# ConnectSecure-API.Part1.ps1 — Dot-sourced by ConnectSecure-API.ps1 (do not run directly)
# Configuration, request pipeline, vulnerability queries, company review endpoints map, through Test-ExternalSubnetConfig.
# --- Configuration ---
$script:ConnectSecureConfig = @{
    BaseUrl = $null
    AccessToken = $null
    TokenExpiry = $null
    TenantName = $null
    ClientId = $null
    ClientSecret = $null
    UserId = $null
    ManualVerificationLogged = $false
    RateLimit = @{
        RequestsPerMinute = 300
        RequestsPerHour = 2000
        RequestsPerDay = 30000
    }
    RequestHistory = [System.Collections.ArrayList]::new()
}

# --- Helper Functions ---

function Remove-SensitiveDataFromObject {
    param([object]$Obj)
    if ($null -eq $Obj) { return $null }
    try {
        $json = $Obj | ConvertTo-Json -Depth 5
        $sensitiveKeys = @('access_token', 'token', 'refresh_token', 'user_id', 'client_secret', 'Client-Auth-Token', 'Authorization')
        foreach ($key in $sensitiveKeys) {
            $json = $json -replace "`"$key`"\s*:\s*`"[^`"]*`"", "`"$key`":`"***REDACTED***`""
            $json = $json -replace "`"$key`"\s*:\s*null", "`"$key`":null"
        }
        return $json
    } catch { return "[Unable to serialize]" }
}

function Write-CSApiLog {
    param(
        [string]$Message,
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [ConnectSecure] [$Level] $Message"
    
    switch ($Level) {
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}

function Test-RateLimit {
    param(
        [int]$MaxRequests,
        [int]$TimeWindowSeconds
    )

    $now = Get-Date
    $windowStart = $now.AddSeconds(-$TimeWindowSeconds)
    
    # Remove old entries outside the time window
    $script:ConnectSecureConfig.RequestHistory = $script:ConnectSecureConfig.RequestHistory | Where-Object { $_ -gt $windowStart }
    $recentRequests = $script:ConnectSecureConfig.RequestHistory.Count
    
    if ($recentRequests -ge $MaxRequests) {
        return $false
    }
    
    return $true
}

function Add-RequestToHistory {
    if ($null -ne $script:ConnectSecureConfig -and $null -ne $script:ConnectSecureConfig.RequestHistory) {
        $null = $script:ConnectSecureConfig.RequestHistory.Add((Get-Date))
    }
}

function Wait-ForRateLimit {
    param(
        [int]$MaxRequests,
        [int]$TimeWindowSeconds
    )

    while (-not (Test-RateLimit -MaxRequests $MaxRequests -TimeWindowSeconds $TimeWindowSeconds)) {
        $waitTime = $TimeWindowSeconds
        Write-CSApiLog "Rate limit reached. Waiting $waitTime seconds..." -Level Warning
        Start-Sleep -Seconds $waitTime
    }
}

function Invoke-ConnectSecureRequest {
    param(
        [string]$Endpoint,
        [string]$Method = "GET",
        [hashtable]$QueryParameters = @{},
        [object]$Body = $null,
        [int]$RetryCount = 3
    )

    # Check rate limits before making request
    Wait-ForRateLimit -MaxRequests $script:ConnectSecureConfig.RateLimit.RequestsPerMinute -TimeWindowSeconds 60
    
    # Ensure we have a valid access token
    if ([string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.AccessToken) -or 
        ($script:ConnectSecureConfig.TokenExpiry -and (Get-Date) -ge $script:ConnectSecureConfig.TokenExpiry)) {
        Write-CSApiLog "Access token expired or missing. Refreshing..." -Level Warning
        Connect-ConnectSecureAPI -BaseUrl $script:ConnectSecureConfig.BaseUrl `
                                  -TenantName $script:ConnectSecureConfig.TenantName `
                                  -ClientId $script:ConnectSecureConfig.ClientId `
                                  -ClientSecret $script:ConnectSecureConfig.ClientSecret
    }

    $url = "$($script:ConnectSecureConfig.BaseUrl)$Endpoint"
    
    # Add query parameters (ensure values are strings for UrlEncode to avoid invocation issues)
    if ($QueryParameters.Count -gt 0) {
        $pairs = @()
        foreach ($kv in $QueryParameters.GetEnumerator()) {
            $valStr = if ($null -eq $kv.Value) { '' } else { $kv.Value.ToString() }
            $pairs += $kv.Key + "=" + [System.Web.HttpUtility]::UrlEncode($valStr)
        }
        $url += "?" + ($pairs -join "&")
    }

    $attempt = 0
    while ($attempt -lt $RetryCount) {
        # Rebuild headers each attempt so we use the latest token (important after 401 refresh)
        # API requires both Authorization and X-USER-ID (per ConnectSecure V4 / Python test)
        $headers = @{
            "Authorization" = "Bearer $($script:ConnectSecureConfig.AccessToken)"
            "Content-Type" = "application/json"
        }
        if (-not [string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.UserId)) {
            $headers['X-USER-ID'] = $script:ConnectSecureConfig.UserId.ToString()
        }

        try {
            Add-RequestToHistory
            
            # Use Invoke-WebRequest instead of Invoke-RestMethod to avoid .NET response parsing 
            # that can throw "null-valued expression" in some PowerShell configurations
            $params = @{
                Uri = $url
                Method = $Method
                Headers = $headers
                UseBasicParsing = $true
                ErrorAction = "Stop"
            }
            if ($Method -ne "GET" -and $null -ne $Body) {
                $params.Body = ($Body | ConvertTo-Json -Depth 10)
            }
            
            $webResponse = Invoke-WebRequest @params
            if ([string]::IsNullOrWhiteSpace($webResponse.Content)) {
                return @{ status = $true; data = @() }
            }
            $response = $webResponse.Content | ConvertFrom-Json
            return $response

        } catch {
            $attempt++
            $errRecord = $_
            $statusCode = $null
            try {
                if ($null -ne $errRecord.Exception.Response -and $null -ne $errRecord.Exception.Response.StatusCode) {
                    $statusCode = $errRecord.Exception.Response.StatusCode.value__
                }
            } catch {
                # StatusCode access can fail with "null-valued expression" in some PowerShell/.NET configurations
            }
            
            if ($statusCode -eq 429) {
                Write-CSApiLog "Rate limit exceeded (429). Waiting before retry..." -Level Warning
                Start-Sleep -Seconds 60
                continue
            } elseif ($statusCode -eq 401) {
                # Log 401 response body for debugging (API may explain why token rejected)
                try {
                    $respStream = $errRecord.Exception.Response.GetResponseStream()
                    if ($null -ne $respStream) {
                        $reader = New-Object System.IO.StreamReader($respStream)
                        $body = $reader.ReadToEnd()
                        $reader.Close()
                        $safeBody = Remove-SensitiveDataFromObject ($body | ConvertFrom-Json -ErrorAction SilentlyContinue)
                        if ($safeBody) { Write-CSApiLog "401 response: $safeBody" -Level Warning } else { Write-CSApiLog "401 received (response body not JSON)" -Level Warning }
                    }
                } catch { }
                Write-CSApiLog "Unauthorized (401). Token invalid or expired." -Level Warning
                Write-CSApiLog "Refreshing token..." -Level Warning
                Connect-ConnectSecureAPI -BaseUrl $script:ConnectSecureConfig.BaseUrl `
                                          -TenantName $script:ConnectSecureConfig.TenantName `
                                          -ClientId $script:ConnectSecureConfig.ClientId `
                                          -ClientSecret $script:ConnectSecureConfig.ClientSecret
                continue
            } elseif ($statusCode -eq 502) {
                Write-CSApiLog "Gateway error (502). Retrying..." -Level Warning
                if ($attempt -lt $RetryCount) {
                    Start-Sleep -Seconds 5
                    continue
                }
            }
            
            if ($attempt -ge $RetryCount) {
                if ($null -eq $statusCode) {
                    Write-CSApiLog "No HTTP response received (connection/network error?). Error: $($errRecord.Exception.Message)" -Level Error
                    if ($errRecord.Exception.Message -like "*null-valued*") {
                        Write-CSApiLog "Stack trace: $($errRecord.ScriptStackTrace)" -Level Error
                        if ($errRecord.Exception.InnerException) {
                            Write-CSApiLog "Inner exception: $($errRecord.Exception.InnerException.Message)" -Level Error
                        }
                    }
                } else {
                    Write-CSApiLog "Request failed after $RetryCount attempts: $($errRecord.Exception.Message)" -Level Error
                }
                throw $errRecord
            }
        }
    }
}

# --- Public Functions ---

function Connect-ConnectSecureAPI {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$TenantName,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientId,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )

    Write-CSApiLog "Connecting to ConnectSecure API..." -Level Info

    # Store configuration
    $script:ConnectSecureConfig.BaseUrl = $BaseUrl.TrimEnd('/')
    $script:ConnectSecureConfig.TenantName = $TenantName
    $script:ConnectSecureConfig.ClientId = $ClientId
    $script:ConnectSecureConfig.ClientSecret = $ClientSecret

    # Create base64 encoded auth token
    # Format: tenant+client_id:client_secret
    # IMPORTANT: Must use UTF8 encoding, not Unicode (causes 502 Gateway Error)
    
    # Validate format
    if ([string]::IsNullOrWhiteSpace($TenantName) -or 
        [string]::IsNullOrWhiteSpace($ClientId) -or 
        [string]::IsNullOrWhiteSpace($ClientSecret)) {
        Write-CSApiLog "Error: Tenant Name, Client ID, and Client Secret are required" -Level Error
        return $false
    }
    
    # Trim whitespace from all inputs
    $TenantName = $TenantName.Trim()
    $ClientId = $ClientId.Trim()
    $ClientSecret = $ClientSecret.Trim()
    
    # Remove any newlines or carriage returns that might have been accidentally copied
    # (but preserve the actual secret value - some secrets can be very long)
    $ClientSecret = $ClientSecret -replace "`r`n|`r|`n", ""
    $ClientId = $ClientId -replace "`r`n|`r|`n", ""
    $TenantName = $TenantName -replace "`r`n|`r|`n", ""
    
    # Construct auth string: tenant+client_id:client_secret
    # Format matches official documentation exactly: tenantname+Client_id:client_secret
    # Use ${} syntax to properly delimit variables when using : separator
    $authString = "${TenantName}+${ClientId}:${ClientSecret}"
    
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
        $base64Auth = [System.Convert]::ToBase64String($bytes)
        Write-CSApiLog "Auth encoding successful" -Level Info
    } catch {
        Write-CSApiLog "Error encoding auth string: $($_.Exception.Message)" -Level Error
        return $false
    }

    try {
        # Match official example exactly
        # Note: Some API implementations require Content-Type even with empty body
        $headers = @{
            "accept" = "application/json"
            "Content-Type" = "application/json"
            "Client-Auth-Token" = $base64Auth
        }

        $authUrl = "$($script:ConnectSecureConfig.BaseUrl)/w/authorize"
        Write-CSApiLog "Authenticating at: $authUrl" -Level Info

        $authUri = [System.Uri]$authUrl

        $maxAuthRetries = 3
        $response = $null
        $apiErrorRetries = 3  # For "Failed to create customer" - known intermittent ConnectSecure issue
        $lastError = $null

        for ($outerAttempt = 1; $outerAttempt -le $apiErrorRetries; $outerAttempt++) {
            for ($attempt = 1; $attempt -le $maxAuthRetries; $attempt++) {
                try {
                    $response = Invoke-RestMethod -Uri $authUri -Method Post -Headers $headers -Body "{}" -TimeoutSec 90 -ErrorAction Stop
                    break
                } catch {
                    $statusCode = $null
                    try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { }
                    $errorDetails = $null
                    try {
                        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                        $errorDetails = $reader.ReadToEnd()
                        $reader.Close()
                    } catch { }

                    Write-CSApiLog "HTTP Error: Status Code $statusCode" -Level Error
                    if ($errorDetails) {
                        $safeErr = try { $parsed = $errorDetails | ConvertFrom-Json; Remove-SensitiveDataFromObject $parsed } catch { $errorDetails -replace '"access_token"\s*:\s*"[^"]*"', '"access_token":"***"' }
                        Write-CSApiLog "Response: $safeErr" -Level Error
                    }

                    if ($statusCode -eq 502) {
                        Write-CSApiLog "502 Gateway Error - often encoding or credentials format. Verify: tenant+client_id:client_secret" -Level Error
                    }
                    if ($statusCode -eq 504) {
                        Write-CSApiLog "504 Gateway Timeout - ConnectSecure server took too long to respond (transient)." -Level Warning
                        if ($attempt -lt $maxAuthRetries) {
                            Write-CSApiLog "Retrying in 10 seconds (attempt $attempt of $maxAuthRetries)..." -Level Info
                            Start-Sleep -Seconds 10
                            continue
                        }
                    }
                    if ($statusCode -eq 502 -and $attempt -lt $maxAuthRetries) {
                        Write-CSApiLog "Retrying in 5 seconds (attempt $attempt of $maxAuthRetries)..." -Level Info
                        Start-Sleep -Seconds 5
                        continue
                    }
                    throw
                }
            }

            Write-CSApiLog "Response received" -Level Info

            # Check for error response first
            if ($response.status -eq $false -or ($response.message -and $response.message.Length -gt 0)) {
            $errorMsg = if ($response.message) { $response.message } else { "Authentication failed" }
            
            # Check if this is a different type of error (not auth failure)
                if ($errorMsg -ne "Failed to authorize") {
                Write-CSApiLog "API returned error message: $errorMsg" -Level Warning
                Write-CSApiLog "This is different from Failed to authorize - credentials appear correct but API returned an error." -Level Warning
                
                # "Failed to tenant info" - tenant lookup failed on ConnectSecure side
                if ($errorMsg -eq "Failed to tenant info") {
                    Write-CSApiLog "" -Level Warning
                    Write-CSApiLog "TENANT LOOKUP FAILED - ConnectSecure could not find your tenant." -Level Warning
                    Write-CSApiLog "  1. Base URL: Use the exact URL from API Key page. Try tenant URL if pod fails: https://YOURTENANT.myconnectsecure.com" -Level Warning
                    Write-CSApiLog "  2. Tenant Name: Must match exactly - case-sensitive, no spaces. Use the subdomain you use to log in." -Level Warning
                    Write-CSApiLog "  3. Pod change: If your tenant moved pods, the old URL may no longer work. Re-copy from ConnectSecure portal." -Level Warning
                    Write-CSApiLog "  4. Re-save: Open API Settings, re-enter credentials (avoid copy-paste artifacts), Save, then Test." -Level Warning
                    Write-CSApiLog "  5. Contact ConnectSecure support if settings are correct - this can be a server-side tenant config issue." -Level Warning
                    Write-CSApiLog "" -Level Warning
                }
                
                # "Failed to create customer" is unexpected for a simple auth test
                if ($errorMsg -eq "Failed to create customer") {
                    Write-CSApiLog "" -Level Warning
                    Write-CSApiLog "UNEXPECTED ERROR: Failed to create customer during authentication" -Level Warning
                    Write-CSApiLog "   This error is unusual for a simple credential test." -Level Warning
                    Write-CSApiLog "   The /w/authorize endpoint should only authenticate, not create customers." -Level Warning
                    Write-CSApiLog "" -Level Warning
                    Write-CSApiLog "   Possible causes:" -Level Warning
                    Write-CSApiLog "   1. ConnectSecure API bug/misconfiguration" -Level Warning
                    Write-CSApiLog "   2. API endpoint trying to auto-provision user account (unexpected behavior)" -Level Warning
                    Write-CSApiLog "   3. API key permissions issue" -Level Warning
                    Write-CSApiLog "   4. Account/tenant configuration issue on ConnectSecure side" -Level Warning
                    Write-CSApiLog "" -Level Warning
                    Write-CSApiLog "   Recommendation: Contact ConnectSecure support with this error message." -Level Warning
                    Write-CSApiLog "   Your credentials appear correct (past Failed to authorize check)." -Level Warning
                    Write-CSApiLog "" -Level Warning
                }
                
                # Check for access_token in various possible fields
                $accessToken = $null
                $userId = $null
                
                if ($response.access_token) {
                    $accessToken = $response.access_token
                    $userId = $response.user_id
                } elseif ($response.token) {
                    $accessToken = $response.token
                    $userId = $response.user_id
                } elseif ($response.data -and $response.data.access_token) {
                    $accessToken = $response.data.access_token
                    $userId = $response.data.user_id
                }
                
                if ($accessToken) {
                    Write-CSApiLog "Found access token in response despite error message - proceeding with authentication" -Level Info
                    $script:ConnectSecureConfig.AccessToken = $accessToken
                    $script:ConnectSecureConfig.UserId = $userId
                    $script:ConnectSecureConfig.TokenExpiry = (Get-Date).AddHours(1)
                    Write-CSApiLog "Successfully authenticated" -Level Success
                    return $true
                }
                # Retry on "Failed to create customer" - known intermittent ConnectSecure API issue
                if ($errorMsg -eq "Failed to create customer" -and $outerAttempt -lt $apiErrorRetries) {
                    Write-CSApiLog "Retrying auth in 8 seconds (attempt $outerAttempt of $apiErrorRetries - this error often succeeds on retry)..." -Level Info
                    Start-Sleep -Seconds 8
                    continue
                }
            }
            
            Write-CSApiLog "Authentication failed: $errorMsg" -Level Error
            Write-CSApiLog "Troubleshooting: Verify Base URL, Tenant Name, and Client ID on ConnectSecure API Key page (Global > Settings > Users > API Key). Ensure API key is Active." -Level Error
            return $false
        }
        
        # Accept both top-level and nested data (API may return { data: { access_token, user_id } })
        $accessToken = $null
        $userId = $null
        if ($response.access_token) {
            $accessToken = $response.access_token
            $userId = $response.user_id
        } elseif ($response.data -and $response.data.access_token) {
            $accessToken = $response.data.access_token
            $userId = $response.data.user_id
        }
        if ($accessToken) {
            $script:ConnectSecureConfig.AccessToken = $accessToken
            $script:ConnectSecureConfig.UserId = $userId
            
            # Set token expiry to 1 hour from now (tokens typically expire after some time)
            $script:ConnectSecureConfig.TokenExpiry = (Get-Date).AddHours(1)
            
            Write-CSApiLog "Successfully authenticated" -Level Success
            return $true
        } else {
            Write-CSApiLog "Authentication failed: No access_token in response" -Level Error
            Write-CSApiLog "Response content: $($response | ConvertTo-Json -Depth 3)" -Level Error
            
            # Common issues
            Write-CSApiLog "Common issues:" -Level Error
            Write-CSApiLog "  1. Check tenant name is correct (no spaces, exact match, case-sensitive)" -Level Error
            Write-CSApiLog "  2. Verify Client ID and Client Secret are correct (copy exactly from ConnectSecure)" -Level Error
            Write-CSApiLog "  3. Ensure format is: tenant+client_id:client_secret (with + separator)" -Level Error
            Write-CSApiLog "  4. Check Base URL is correct - e.g. https://pod0.myconnectsecure.com" -Level Error
            Write-CSApiLog "  5. Verify your API key is active in ConnectSecure portal" -Level Error
            
            return $false
        }

        }  # end for ($outerAttempt)

    } catch {
        Write-CSApiLog "Authentication failed: $($_.Exception.Message)" -Level Error
        if ($_.ErrorDetails.Message) {
            Write-CSApiLog "Error details: $($_.ErrorDetails.Message)" -Level Error
        }
        return $false
    }
}

function Get-ConnectSecureCompanyDisplayInfo {
    param([object]$Company)
    # Skip non-objects (e.g. boolean true from API)
    if ($null -eq $Company -or ($Company -is [bool])) {
        return @{ Name = ""; Id = "" }
    }
    # ConnectSecure/Elasticsearch may use various property names; try all known variants
    $name = ""
    $id = ""
    
    # Direct properties
    $nameProps = @("name","company_name","companyName","title","label","displayName","customer_name")
    foreach ($p in $nameProps) {
        if ($Company.PSObject.Properties[$p] -and -not [string]::IsNullOrWhiteSpace($Company.$p)) {
            $name = $Company.$p
            break
        }
    }
    
    # Nested _source (Elasticsearch style)
    if ([string]::IsNullOrWhiteSpace($name) -and $Company._source) {
        foreach ($p in $nameProps) {
            if ($Company._source.PSObject.Properties[$p] -and -not [string]::IsNullOrWhiteSpace($Company._source.$p)) {
                $name = $Company._source.$p
                break
            }
        }
    }
    
    # name.keyword (Elasticsearch keyword field)
    if ([string]::IsNullOrWhiteSpace($name) -and $Company['name.keyword']) { $name = $Company['name.keyword'] }
    if ([string]::IsNullOrWhiteSpace($name) -and $Company._source['name.keyword']) { $name = $Company._source['name.keyword'] }
    
    # Fallback: any property containing name with a non-empty value
    if ([string]::IsNullOrWhiteSpace($name)) {
        foreach ($prop in $Company.PSObject.Properties) {
            if ($prop.Name -match 'name' -and $prop.Value -and $prop.Value -is [string] -and $prop.Value.Trim()) {
                $name = $prop.Value
                break
            }
        }
    }
    
    # ID - direct properties
    $idProps = @('id','company_id','companyId','_id')
    foreach ($p in $idProps) {
        if ($Company.PSObject.Properties[$p] -and $null -ne $Company.$p) {
            $id = $Company.$p
            break
        }
    }
    
    # Nested _source
    if ($null -eq $id -or $id -eq '') {
        if ($Company._source) {
            foreach ($p in $idProps) {
                if ($Company._source.PSObject.Properties[$p] -and $null -ne $Company._source.$p) {
                    $id = $Company._source.$p
                    break
                }
            }
        }
    }
    
    # Fallback: any property containing id with a value (exclude guid-like user_id for name)
    if ($null -eq $id -or $id -eq '') {
        foreach ($prop in $Company.PSObject.Properties) {
            if (($prop.Name -eq 'id' -or $prop.Name -eq 'company_id' -or $prop.Name -eq '_id') -and $null -ne $prop.Value) {
                $id = $prop.Value
                break
            }
        }
    }
    
    return @{ Name = $name; Id = $id }
}

function Get-ConnectSecureCompanies {
    param(
        [int]$Limit = 1000,
        [int]$Skip = 0,
        [switch]$FetchAll
    )

    Write-CSApiLog ('Fetching companies (Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info

    $allCompanies = @()
    $currentSkip = $Skip
    $pageSize = [Math]::Min($Limit, 1000)

    do {
        $queryParams = @{
            limit = $pageSize
            skip = $currentSkip
        }

        try {
            $response = Invoke-ConnectSecureRequest -Endpoint '/r/company/companies' -QueryParameters $queryParams
            
            if ($null -eq $response) {
                Write-CSApiLog 'No company data returned (request failed or returned null)' -Level Warning
                break
            }
            
            # Handle different API response formats (list, object with data, or Elasticsearch hits)
            $companies = @()
            if ($response -is [array]) {
                $companies = $response
            } elseif ($response.data) {
                $companies = $response.data
            } elseif ($response.hits -and $response.hits.hits) {
                # Elasticsearch format: use _source from each hit
                $companies = $response.hits.hits | ForEach-Object { if ($_._source) { $_._source } else { $_ } }
            } elseif ($response.status -and $response.data) {
                $companies = $response.data
            }
            
            if ($companies.Count -gt 0) {
                # ConnectSecure may return [true, true, true] - array of booleans, not company objects
                $first = $companies[0]
                if ($first -is [bool]) {
                    Write-CSApiLog 'API returned boolean array (minimal data). Trying company_stats for company names...' -Level Warning
                    $allCompanies = @()
                    try {
                        $statsResponse = Invoke-ConnectSecureRequest -Endpoint '/r/company/company_stats' -QueryParameters @{ limit = 1000; skip = 0 }
                        $stats = @()
                        if ($statsResponse.data) { $stats = $statsResponse.data }
                        elseif ($statsResponse -is [array]) { $stats = $statsResponse }
                        $seenIds = @{}
                        foreach ($s in $stats) {
                            if ($s -is [bool]) { continue }
                            $cid = $s.company_id; if ($null -eq $cid) { $cid = $s.companyId }; if ($null -eq $cid) { $cid = $s.id }
                            if ($null -ne $cid -and -not $seenIds[$cid]) {
                                $seenIds[$cid] = $true
                                $cname = $s.company_name; if ($null -eq $cname) { $cname = $s.companyName }; if ($null -eq $cname) { $cname = $s.name }; if ($null -eq $cname) { $cname = $s.title }; if ($null -eq $cname) { $cname = "" }
                                $allCompanies += [PSCustomObject]@{ id = $cid; name = $cname }
                            }
                        }
                        Write-CSApiLog ('Got ' + $allCompanies.Count + ' companies from company_stats') -Level Info
                    } catch {
                        Write-CSApiLog ('company_stats fallback failed: ' + $_.Exception.Message) -Level Warning
                    }
                    # If we still have no companies (company_stats also returned booleans or failed), use placeholder IDs from count
                    if ($allCompanies.Count -eq 0) {
                        Write-CSApiLog 'Using placeholder company entries - API returns minimal data, consider contacting ConnectSecure support for full company list' -Level Warning
                        for ($i = 0; $i -lt $companies.Count; $i++) {
                            $allCompanies += [PSCustomObject]@{ id = $i + 1; name = "" }
                        }
                    }
                    break
                }
                
                $allCompanies += $companies
                Write-CSApiLog ('Retrieved ' + $companies.Count + ' companies (Total: ' + $allCompanies.Count + ')') -Level Info
                
                # Pagination: CyberCNS uses skip=page index (0, 1, 2...). Continue if FetchAll and got full page.
                if ($FetchAll -and $companies.Count -eq $pageSize) {
                    $currentSkip += 1
                } else {
                    break
                }
            } else {
                break
            }
        } catch {
            Write-CSApiLog ('Error fetching companies: ' + $_.Exception.Message) -Level Error
            throw
        }
    } while ($FetchAll)

    if ($allCompanies.Count -gt 0) {
        Write-CSApiLog ('Retrieved ' + $allCompanies.Count + ' companies') -Level Success
        return $allCompanies
    } else {
        Write-CSApiLog 'No company data returned' -Level Warning
        return @()
    }
}

# Vulnerability report_queries endpoints - used by Get-ConnectSecure*Vulnerabilities
# vulnerabilities_details: one row per unique vuln with host_name (semicolon list), affected_assets (compact, ~10-50k vs 200k+)
# asset_wise_vulnerabilities: one row per host+vuln (use for per-host detail if needed)
$script:VulnEndpoints = @{
    'application'  = '/r/report_queries/vulnerabilities_details'
    'external'     = '/r/report_queries/external_asset_vulnerabilities'
    'suppressed'   = '/r/report_queries/vulnerabilities_details_suppressed'
    'registry'     = '/r/report_queries/registry_problems_remediation'
    'network'      = '/r/report_queries/application_vulnerabilities_net'
}
# Max records to fetch (0 = unlimited). Stops after this - ~5 pages at 5k each. Set higher if you need more.
$script:VulnMaxRecords = 25000
# When using asset_wise_vulnerabilities: aggregate to one row per unique vuln with AffectedHosts (reduces 200k→~50k). Ignored for vulnerabilities_details.
$script:VulnAggregateByVulnerability = $true
# API filter - ConnectSecure may ignore this. Set to empty string to request all severities from API.
$script:VulnSeverityFilter = ''
# Client-side filter: after download, keep only Critical and High (API filter often ignored). Set $false to include all severities (Critical, High, Medium, Low).
$script:VulnFilterCriticalHighOnly = $false
# Try server-side company filter via condition param (API may support it; tested formats caused 400/errors - keep false for now)
$script:UseConditionForCompanyFilter = $false

function Test-RowMatchesCompanyId {
    <#
    .SYNOPSIS
    Returns $true if the row's company_ids or company_id contains the given CompanyId.
    Handles company_ids, companyIds, company_id, _source nesting, string/array formats.
    #>
    param([object]$Row, [string]$CompanyIdStr)
    if ([string]::IsNullOrWhiteSpace($CompanyIdStr)) { return $true }
    $obj = $Row
    if ($null -ne $Row._source) { $obj = $Row._source }
    if ($null -eq $obj) { return $false }
    $cids = $obj.company_ids
    if ($null -eq $cids) { $cids = $obj.companyIds }
    if ($null -eq $cids) { $cids = $obj.company_id }
    if ($null -eq $cids) { return $false }
    $cidsStr = if ($cids -is [array]) { ($cids | ForEach-Object { [string]$_ }) -join ';' } else { [string]$cids }
    return $cidsStr -match ('(^|;)' + [regex]::Escape($CompanyIdStr) + '($|;)')
}

function Invoke-ConnectSecureVulnerabilityQuery {
    <#
    .SYNOPSIS
    Shared implementation for fetching vulnerability data from report_queries endpoints.
    .PARAMETER VulnType
    One of: application, external, suppressed, registry, network
    .PARAMETER MaxRecords
    Override script VulnMaxRecords for this call. 0 = use script default.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('application', 'external', 'suppressed', 'registry', 'network')]
        [string]$VulnType,
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [bool]$FetchAll = $false,
        [int]$MaxRecords = -1,
        [switch]$RegistryRemediatedOnly,
        [switch]$RegistrySuppressedOnly
    )

    $endpoint = $script:VulnEndpoints[$VulnType]
    $label = switch ($VulnType) {
        'application' { 'vulnerabilities' }
        'external'    { 'external vulnerabilities' }
        'suppressed'  { 'suppressed vulnerabilities' }
        'registry'    { 'registry vulnerabilities' }
        'network'     { 'network vulnerabilities' }
        default       { 'vulnerabilities' }
    }

    Write-CSApiLog ('Fetching ' + $label + ' (CompanyId: ' + $CompanyId + ', Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info

    $allData = @()
    $currentSkip = $Skip
    $pageSize = $Limit
    $maxPages = 200  # Safety: 200 x 5000 = 1M records max

    do {
        $queryParams = @{ limit = $pageSize; skip = $currentSkip }
        # Registry and network use condition+order_by (per portal capture); others use company_id+sort
        if ($VulnType -eq 'registry') {
            $queryParams.order_by = 'severity asc'
            if ($CompanyId -gt 0) {
                $registryClause = if ($RegistryRemediatedOnly) { 'is_remediated = true' }
                    elseif ($RegistrySuppressedOnly) { 'is_suppressed=true and is_remediated = false' }
                    else { 'is_suppressed=false and is_remediated = false' }
                $queryParams.condition = "company_id=$CompanyId and $registryClause"
            }
        } elseif ($VulnType -eq 'network') {
            $queryParams.order_by = 'affected_assets desc'
            # API requires condition to return network vulns; without it the report can be blank
            $netCondition = "software_type='networksoftware' and unconfirmed = 'false'"
            if ($CompanyId -gt 0) {
                $queryParams.condition = "company_id=$CompanyId and $netCondition"
            } else {
                $queryParams.condition = $netCondition
            }
        } else {
            $queryParams.sort = $Sort
            if ($CompanyId -gt 0) {
                $queryParams.company_id = $CompanyId
                if ($script:UseConditionForCompanyFilter) {
                    $queryParams.condition = 'company_ids:' + $CompanyId
                }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($Filter)) { $queryParams.filter = $Filter }

        try {
            $response = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters $queryParams

            # Extract data: standard { status, data } or Elasticsearch { hits: { hits: [...] } }
            $pageData = $null
            if ($response.status -and $response.data) {
                $pageData = $response.data
            } elseif ($response.hits -and $response.hits.hits) {
                $pageData = $response.hits.hits | ForEach-Object { if ($_._source) { $_._source } else { $_ } }
                Write-CSApiLog ('API returned Elasticsearch format: extracted ' + $pageData.Count + ' rows') -Level Info
            }

            if ($pageData -and $pageData.Count -gt 0) {
                $allData += $pageData
                $pageNum = [Math]::Floor($currentSkip / $pageSize) + 1
                Write-CSApiLog ('Page ' + $pageNum + ' : retrieved ' + $pageData.Count + ' ' + $label + ' (Total: ' + $allData.Count + ')') -Level Info

                $maxRec = if ($MaxRecords -ge 0) { $MaxRecords } else { $script:VulnMaxRecords }
                if ($maxRec -gt 0 -and $allData.Count -ge $maxRec) {
                    Write-CSApiLog ('Reached record limit (' + $maxRec + ').') -Level Warning
                    break
                }
                if ($FetchAll -eq $true -and $pageData.Count -eq $pageSize) {
                    if ($pageNum -ge $maxPages) {
                        Write-CSApiLog ('Reached safety limit ' + $maxPages + ' pages. Use higher Limit or paginate manually for more.') -Level Warning
                        break
                    }
                    $currentSkip += $pageSize
                } else {
                    break
                }
            } else {
                Write-CSApiLog ('No more ' + $label + ' data returned') -Level Info
                break
            }
        } catch {
            Write-CSApiLog ('Error fetching ' + $label + ' : ' + $_.Exception.Message) -Level Error
            throw
        }
    } while ($FetchAll)

    Write-CSApiLog ('Total ' + $label + ' retrieved: ' + $allData.Count) -Level Success
    return $allData
}

function ConvertFrom-VulnerabilitiesDetailsFormat {
    <#
    .SYNOPSIS
    Normalizes vulnerabilities_details API response (one row per vuln, host_name semicolon list) to _aggregated report format.
    #>
    param([array]$Data)
    if (-not $Data -or $Data.Count -eq 0) { return @() }
    $result = @()
    foreach ($r in $Data) {
        $hostStr = $r.host_name
        if (-not $hostStr -and $r.Host_Name) { $hostStr = $r.Host_Name }
        $hostStr = if ($hostStr) { ($hostStr -as [string]).Trim() } else { "" }
        $count = $r.affected_assets
        if ($null -eq $count -or $count -eq '') { $count = 0 }
        if ($hostStr) { $count = [Math]::Max($count, ($hostStr -split ';').Count) }
        $result += [PSCustomObject]@{
            problem_id = $r.problem_id
            problem_name = $r.problem_name
            software_name = if ($r.software_name) { $r.software_name } elseif ($r.problem_name) { [string]$r.problem_name } else { "" }
            severity = $r.severity
            description = $r.description
            epss_score = $r.epss_score
            base_score = $r.base_score
            AffectedHosts = $hostStr
            HostCount = [int]$count
            _aggregated = $true
        }
    }
    Write-CSApiLog ('Normalized ' + $Data.Count + ' vuln-details rows to report format') -Level Info
    return $result
}

function Invoke-AggregateVulnerabilitiesByUniqueVuln {
    <#
    .SYNOPSIS
    Aggregates asset_wise rows (one per host+vuln) into one row per unique vulnerability with AffectedHosts list.
    Reduces 200k redundant rows to ~10-50k unique vulns.
    #>
    param([array]$AssetWiseData)
    if (-not $AssetWiseData -or $AssetWiseData.Count -eq 0) { return @() }
    $grouped = $AssetWiseData | Group-Object -Property {
        $vulnId = $_.problem_id; $prod = $_.software_name; $sev = $_.severity
        if (-not $vulnId) { $vulnId = $_.problem_name }
        ($vulnId, $prod, $sev) -join '|'
    }
    $result = @()
    foreach ($g in $grouped) {
        $first = $g.Group[0]
        $hosts = $g.Group | ForEach-Object {
            $h = $_.host_name; if (-not $h) { $h = $_.'Host Name' }
            $i = $_.ip; if (-not $i) { $i = $_.IP }
            if ($h -and $i) { $h + ' (' + $i + ')' } elseif ($h) { $h } elseif ($i) { $i } else { $null }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        $affectedHosts = ($hosts -join '; ')
        $result += [PSCustomObject]@{
            problem_id = $first.problem_id
            problem_name = $first.problem_name
            software_name = $first.software_name
            severity = $first.severity
            description = $first.description
            epss_score = $first.epss_score
            base_score = $first.base_score
            AffectedHosts = $affectedHosts
            HostCount = $hosts.Count
            _aggregated = $true
        }
    }
    Write-CSApiLog ('Aggregated ' + $AssetWiseData.Count + ' rows to ' + $result.Count + ' unique vulnerabilities') -Level Success
    return $result
}

function Get-ConnectSecureVulnerabilities {
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll,
        [switch]$Raw
    )
    $useFilter = if ([string]::IsNullOrWhiteSpace($Filter) -and $script:VulnSeverityFilter) { $script:VulnSeverityFilter } else { $Filter }
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    $data = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'application' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $useFilter -Sort $Sort -FetchAll $doFetchAll
    # Client-side filter by company_ids: API returns tenant-wide data (vulnerabilities_details ignores company_id param per swagger)
    if ($CompanyId -gt 0 -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $cidStr = [string]$CompanyId
        $data = $data | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($data.Count -lt $before) {
            Write-CSApiLog ('Filtered to company ' + $CompanyId + ': ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
        }
    }
    if ($script:VulnFilterCriticalHighOnly -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $data = $data | Where-Object {
            $s = $_.severity; if (-not $s) { $s = $_.Severity }
            $s -eq 'Critical' -or $s -eq 'High'
        }
        Write-CSApiLog ('Filtered to Critical+High: ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
    }
    if ($Raw) { return $data }
    $ep = $script:VulnEndpoints['application']
    if ($ep -match 'vulnerabilities_details') {
        return ConvertFrom-VulnerabilitiesDetailsFormat -Data $data
    }
    if ($script:VulnAggregateByVulnerability -and $data -and $data.Count -gt 0) {
        return Invoke-AggregateVulnerabilitiesByUniqueVuln -AssetWiseData $data
    }
    return $data
}

function Get-ConnectSecureExternalVulnerabilities {
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll
    )
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    $data = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'external' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $Filter -Sort $Sort -FetchAll $doFetchAll
    if ($CompanyId -gt 0 -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $cidStr = [string]$CompanyId
        $data = $data | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($data.Count -lt $before) {
            Write-CSApiLog ('Filtered external to company ' + $CompanyId + ': ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
        }
    }
    return $data
}

function Get-ConnectSecureSuppressedVulnerabilities {
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll,
        [switch]$Raw
    )
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    $data = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'suppressed' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $Filter -Sort $Sort -FetchAll $doFetchAll
    if ($CompanyId -gt 0 -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $cidStr = [string]$CompanyId
        $data = $data | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($data.Count -lt $before) {
            Write-CSApiLog ('Filtered suppressed to company ' + $CompanyId + ': ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
        }
    }
    if ($Raw) { return $data }
    if ($script:VulnEndpoints['suppressed'] -match 'vulnerabilities_details') {
        return ConvertFrom-VulnerabilitiesDetailsFormat -Data $data
    }
    return $data
}

function Get-ConnectSecureRegistryVulnerabilities {
    <#
    .SYNOPSIS
    Fetches registry/misconfiguration vulnerabilities from registry_problems_remediation.
    Returns raw API rows; use ConvertFrom-RegistryProblemsFormat for 13-column Excel export.
    .PARAMETER RemediatedOnly
    When set, fetches remediated issues (is_remediated = true). Default fetches open issues (is_remediated = false).
    .PARAMETER SuppressedOnly
    When set, fetches suppressed issues (is_suppressed=true and is_remediated = false).
    #>
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll,
        [switch]$RemediatedOnly,
        [switch]$SuppressedOnly
    )
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    $data = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'registry' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $Filter -Sort $Sort -FetchAll $doFetchAll -RegistryRemediatedOnly:$RemediatedOnly -RegistrySuppressedOnly:$SuppressedOnly
    if ($CompanyId -gt 0 -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $cidStr = [string]$CompanyId
        $data = $data | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($data.Count -lt $before) {
            Write-CSApiLog ('Filtered registry to company ' + $CompanyId + ': ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
        }
    }
    return $data
}

function Get-ConnectSecureNetworkVulnerabilities {
    <#
    .SYNOPSIS
    Fetches network-scoped application vulnerabilities from application_vulnerabilities_net.
    Returns software-level aggregated data (no per-asset host_name/ip).
    #>
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [switch]$FetchAll
    )
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    $data = Invoke-ConnectSecureVulnerabilityQuery -VulnType 'network' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $Filter -Sort $Sort -FetchAll $doFetchAll
    if ($CompanyId -gt 0 -and $data -and $data.Count -gt 0) {
        $before = $data.Count
        $cidStr = [string]$CompanyId
        $data = $data | Where-Object { Test-RowMatchesCompanyId -Row $_ -CompanyIdStr $cidStr }
        if ($data.Count -lt $before) {
            Write-CSApiLog ('Filtered network to company ' + $CompanyId + ': ' + $before + ' -> ' + $data.Count + ' rows') -Level Info
        }
    }
    return $data
}

function Get-ConnectSecureAssets {
    <#
    .SYNOPSIS
    Fetches assets from ConnectSecure API. Used to build asset id -> name lookup for enriching vulnerability data.
    .PARAMETER CompanyId
    Company ID (0 = all).
    .PARAMETER Limit
    Page size (default 5000).
    .PARAMETER Skip
    Offset for pagination.
    .PARAMETER FetchAll
    If true, paginates until all assets are retrieved.
    #>
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [switch]$FetchAll
    )
    $doFetchAll = if ($FetchAll) { $true } else { $false }
    Write-CSApiLog ('Fetching assets (CompanyId: ' + $CompanyId + ', Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info
    $allData = @()
    $currentSkip = $Skip
    $pageSize = [Math]::Max(100, [Math]::Min(5000, $Limit))
    $maxPages = 100
    $pageNum = 0
    $assetsCondition = $null
    if ($CompanyId -gt 0) {
        if ($script:UseConditionForCompanyFilter) {
            $assetsCondition = 'company_ids:' + $CompanyId
        } else {
            # Try condition=company_id=X (same format as registry/network). Swagger shows assets accepts condition.
            $assetsCondition = 'company_id=' + $CompanyId
        }
    }
    do {
        $queryParams = @{ limit = $pageSize; skip = $currentSkip }
        if ($assetsCondition) { $queryParams.condition = $assetsCondition }
        try {
            $response = Invoke-ConnectSecureRequest -Endpoint '/r/asset/assets' -QueryParameters $queryParams
            $assets = @()
            if ($response -is [array]) { $assets = $response }
            elseif ($response.data) { $assets = $response.data }
            elseif ($response.status -and $response.data) { $assets = $response.data }
            if ($assets.Count -gt 0) {
                $allData += $assets
                $pageNum++
                Write-CSApiLog ('Assets page ' + $pageNum + ': retrieved ' + $assets.Count + ' (Total: ' + $allData.Count + ')') -Level Info
                if ($doFetchAll -and $assets.Count -eq $pageSize -and $pageNum -lt $maxPages) {
                    $currentSkip += $pageSize
                } else { break }
            } else { break }
        } catch {
            if ($assetsCondition -and $_.Exception.Message -match '400|Bad Request') {
                Write-CSApiLog ('Assets condition rejected (400), retrying without company filter: ' + $assetsCondition) -Level Warning
                $assetsCondition = $null
                $currentSkip = $Skip
                $allData = @()
                $pageNum = 0
                continue
            }
            Write-CSApiLog ('Error fetching assets: ' + $_.Exception.Message) -Level Error
            throw
        }
    } while ($doFetchAll)
    Write-CSApiLog ('Total assets retrieved: ' + $allData.Count) -Level Success
    return $allData
}

# --- Company Review (agents, probes, credentials, firewall, scan dates) ---
# NOTE: Before adding client-side filtering for any endpoint, capture from the web portal first
# (see api/CAPTURE-PORTAL-COMPANY-REVIEW-GUIDE.md). The portal may use server-side params we can adopt.
# Endpoints from swagger.yaml
$script:CompanyReviewEndpoints = @{
    JobsView                  = '/r/company/jobs_view'
    CompanyStats              = '/r/company/company_stats'
    LightweightAssets         = '/r/report_queries/lightweight_assets'
    Agents                    = '/r/company/agents'
    Credentials               = '/r/company/credentials'
    AgentCredentialsMapping   = '/r/company/agent_credentials_mapping'
    DiscoverySettings         = '/r/company/discovery_settings'
    DiscoverySettingsReport   = '/r/report_queries/discovery_settings'
    AgentDiscoveryMapping     = '/r/company/agent_discoverysettings_mapping'
    AgentDiscoveryCredentials = '/r/company/agent_discovery_credentials'
    AssetFirewallPolicy       = '/r/asset/asset_firewall_policy'
    FirewallAssetView         = '/r/report_queries/firewall_asset_view'
    Assets                    = '/r/asset/assets'
    AssetView                 = '/r/asset/asset_view'
}

function Get-NetworkBroadcastFromCidr {
    <#
    .SYNOPSIS
    Returns network and broadcast addresses for a CIDR. For /29, usable IPs exclude network, broadcast, and typically gateway.
    #>
    param([string]$Cidr)
    if ([string]::IsNullOrWhiteSpace($Cidr) -or $Cidr -notmatch '^(\d+\.\d+\.\d+\.\d+)/(\d+)$') { return $null }
    $ip = $Matches[1]; $prefix = [int]$Matches[2]
    $ipLong = 0
    foreach ($octet in ($ip -split '\.')) { $ipLong = ($ipLong -shl 8) + [int]$octet }
    $mask = if ($prefix -ge 32) { 0xFFFFFFFF } else { [uint32]::MaxValue -shl (32 - $prefix) }
    $networkLong = $ipLong -band $mask
    $broadcastLong = $networkLong -bor (-bnot $mask -band 0xFFFFFFFF)
    $netStr = (($networkLong -shr 24) -band 0xFF).ToString() + '.' + (($networkLong -shr 16) -band 0xFF).ToString() + '.' + (($networkLong -shr 8) -band 0xFF).ToString() + '.' + ($networkLong -band 0xFF).ToString()
    $bcastStr = (($broadcastLong -shr 24) -band 0xFF).ToString() + '.' + (($broadcastLong -shr 16) -band 0xFF).ToString() + '.' + (($broadcastLong -shr 8) -band 0xFF).ToString() + '.' + ($broadcastLong -band 0xFF).ToString()
    return @{ Network = $netStr; Broadcast = $bcastStr; Prefix = $prefix }
}

function Test-ExternalSubnetConfig {
    <#
    .SYNOPSIS
    Checks if scan targets include network/broadcast (and optionally gateway). Returns array of issue strings.
    #>
    param([string[]]$Targets, [string]$Cidr)
    $issues = [System.Collections.ArrayList]::new()
    $nb = Get-NetworkBroadcastFromCidr -Cidr $Cidr
    switch (-not $nb) { $true { return @() } }
    foreach ($t in $Targets) {
        $t = ($t -replace '\s+', '').Trim()
        switch ([string]::IsNullOrWhiteSpace($t)) { $true { continue } }
        switch ($t -eq $nb.Network) { $true { [void]$issues.Add("Target includes network address: $t") } }
        switch ($t -eq $nb.Broadcast) { $true { [void]$issues.Add("Target includes broadcast address: $t") } }
    }
    return @($issues)
}
