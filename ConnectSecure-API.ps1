#Requires -Version 5.1
<#
.SYNOPSIS
ConnectSecure API Client Module

.DESCRIPTION
Provides functions to interact with the ConnectSecure API v4 for fetching vulnerability data.

.NOTES
Version: 1.0.0
Author: River Run MSP
#>

# --- Add Required Assemblies ---
Add-Type -AssemblyName System.Web

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
    
    $recentRequests = ($script:ConnectSecureConfig.RequestHistory | Where-Object { $_ -gt $windowStart }).Count
    
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
    
    # Add query parameters
    if ($QueryParameters.Count -gt 0) {
        $queryString = ($QueryParameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$([System.Web.HttpUtility]::UrlEncode($_.Value))" }) -join "&"
        $url += "?$queryString"
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
                        Write-CSApiLog "401 response body: $body" -Level Warning
                    }
                } catch { }
                $tokenPreview = if ($script:ConnectSecureConfig.AccessToken) { $script:ConnectSecureConfig.AccessToken.Substring(0, [Math]::Min(50, $script:ConnectSecureConfig.AccessToken.Length)) } else { "NULL" }
                Write-CSApiLog "Unauthorized (401). Token sent (first 50 chars): $tokenPreview..." -Level Warning
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
    
    Write-CSApiLog "Client Secret length: $($ClientSecret.Length) characters" -Level Info
    
    # Construct auth string: tenant+client_id:client_secret
    # Format matches official documentation exactly: tenantname+Client_id:client_secret
    # Use ${} syntax to properly delimit variables when using : separator
    $authString = "${TenantName}+${ClientId}:${ClientSecret}"
    
    # Log auth string format (without secret) for debugging
    $authStringPreview = "${TenantName}+${ClientId}:***"
    Write-CSApiLog "Auth string format: $authStringPreview" -Level Info
    Write-CSApiLog "Auth string length: $($authString.Length) characters" -Level Info
    
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
        $base64Auth = [System.Convert]::ToBase64String($bytes)
        Write-CSApiLog "Base64 encoding successful (UTF8)" -Level Info
        Write-CSApiLog "Base64 token length: $($base64Auth.Length) characters" -Level Info
        Write-CSApiLog "Base64 token (first 50 chars): $($base64Auth.Substring(0, [Math]::Min(50, $base64Auth.Length)))..." -Level Info
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
        Write-CSApiLog "Using tenant: $TenantName" -Level Info
        Write-CSApiLog "Headers being sent:" -Level Info
        Write-CSApiLog "  accept: application/json" -Level Info
        Write-CSApiLog "  Content-Type: application/json" -Level Info
        Write-CSApiLog "  Client-Auth-Token: $($base64Auth.Substring(0, [Math]::Min(80, $base64Auth.Length)))..." -Level Info
        Write-CSApiLog "  Client-Auth-Token length: $($base64Auth.Length) characters" -Level Info

        $authUri = [System.Uri]$authUrl
        Write-CSApiLog "Constructed API URL: $authUri" -Level Info
        if (-not $script:ConnectSecureConfig.ManualVerificationLogged) {
            $script:ConnectSecureConfig.ManualVerificationLogged = $true
            Write-CSApiLog "Manual verification: format ${TenantName}+${ClientId}:<secret>, UTF8 base64 at base64encode.org" -Level Info
        }

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
                    if ($errorDetails) { Write-CSApiLog "Response body: $errorDetails" -Level Error }

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

            # Log full response for debugging
            Write-CSApiLog "Response received. Checking for access token..." -Level Info
            if ($response) {
                $responseKeys = $response | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                $keyDelim = ", "
                Write-CSApiLog "Response contains keys: $($responseKeys -join $keyDelim)" -Level Info
                Write-CSApiLog "Full response: $($response | ConvertTo-Json -Depth 3)" -Level Info
            }

            # Check for error response first
            if ($response.status -eq $false -or ($response.message -and $response.message.Length -gt 0)) {
            $errorMsg = if ($response.message) { $response.message } else { "Authentication failed" }
            
            # Check if this is a different type of error (not auth failure)
            if ($errorMsg -ne "Failed to authorize") {
                Write-CSApiLog "API returned error message: $errorMsg" -Level Warning
                Write-CSApiLog "This is different from Failed to authorize - credentials appear correct but API returned an error." -Level Warning
                Write-CSApiLog "Full response: $($response | ConvertTo-Json -Depth 3)" -Level Warning
                
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
                    Write-CSApiLog "Successfully authenticated. User ID: $userId" -Level Success
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
            Write-CSApiLog "Full response: $($response | ConvertTo-Json -Depth 3)" -Level Error
            
            # Show what was sent (for debugging)
            Write-CSApiLog "Debugging info:" -Level Error
            Write-CSApiLog "  Tenant Name: $TenantName (Length: $($TenantName.Length))" -Level Error
            Write-CSApiLog "  Client ID: $ClientId (Length: $($ClientId.Length))" -Level Error
            Write-CSApiLog "  Client Secret: *** (Length: $($ClientSecret.Length) characters)" -Level Error
            Write-CSApiLog "  Base URL: $($script:ConnectSecureConfig.BaseUrl)" -Level Error
            
            # Check if secret might have hidden characters
            if ($ClientSecret -match "`r|`n|`t") {
                Write-CSApiLog "  WARNING: Client Secret contains newlines or tabs - these have been removed" -Level Warning
            }
            
            # Provide specific troubleshooting based on error
            Write-CSApiLog "Troubleshooting steps:" -Level Error
            Write-CSApiLog "  CRITICAL: Verify Base URL and Tenant Name" -Level Error
            Write-CSApiLog "     Base URL: $($script:ConnectSecureConfig.BaseUrl)" -Level Error
            Write-CSApiLog "     Tenant Name: $TenantName" -Level Error
            Write-CSApiLog "     Client ID: $ClientId" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "  🔍 TENANT NAME VERIFICATION (Most Likely Issue):" -Level Error
            Write-CSApiLog "     The tenant name format is critical and must match EXACTLY." -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "     Steps to find correct tenant name:" -Level Error
            Write-CSApiLog "     1. Log into ConnectSecure portal: $($script:ConnectSecureConfig.BaseUrl)" -Level Error
            Write-CSApiLog "     2. Go to Global, Settings, Users, select your user, Action, API Key" -Level Error
            Write-CSApiLog "     3. Look for Tenant Name or Tenant field on that page" -Level Error
            Write-CSApiLog "     4. Copy it EXACTLY as shown - case-sensitive, no spaces" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "     Common tenant name formats:" -Level Error
            Write-CSApiLog "     - Company identifier: river-run, acme-corp (MOST COMMON - found on API Key page)" -Level Error
            Write-CSApiLog "     - Email address: user@company.com (sometimes used)" -Level Error
            Write-CSApiLog "     - UUID/GUID format" -Level Error
            Write-CSApiLog "     - Numeric ID" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "     IMPORTANT: Tenant name is typically the company identifier (e.g. river-run)" -Level Error
            Write-CSApiLog "     IMPORTANT: Found on API Key page as Tenant Name field" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "     Current tenant name: $TenantName" -Level Error
            Write-CSApiLog "     If this does not match what is on the API Key page, that is the problem!" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "     IMPORTANT: The tenant name might be:" -Level Error
            Write-CSApiLog "     - Your email address (not river-run)" -Level Error
            Write-CSApiLog "     - A different format entirely" -Level Error
            Write-CSApiLog "     - Case-sensitive (River-Run vs river-run)" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "  2. VERIFY API KEY STATUS:" -Level Error
            Write-CSApiLog "     - Ensure API key shows as Active (not expired/revoked)" -Level Error
            Write-CSApiLog "     - If you just reset it, wait a few seconds and try again" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "  3. RUN MANUAL TEST SCRIPT:" -Level Error
            $authTest = ('     - Run: .\Test-ConnectSecure-Auth.ps1 -BaseUrl {0} -TenantName {1} -ClientId {2} -ClientSecret your-secret' -f $script:ConnectSecureConfig.BaseUrl, $TenantName, $ClientId)
Write-CSApiLog $authTest -Level Error
            Write-CSApiLog "     - This will test the exact request and show detailed debugging" -Level Error
            Write-CSApiLog "" -Level Error
            Write-CSApiLog "  4. MANUAL BASE64 TEST:" -Level Error
            Write-CSApiLog "     - Use the base64 token shown above in Postman/curl" -Level Error
            Write-CSApiLog "     - Or regenerate at: https://www.base64encode.org/" -Level Error
            Write-CSApiLog "     - Format: ${TenantName}+${ClientId}:<your-secret>" -Level Error
            Write-CSApiLog "     - Select UTF8 encoding" -Level Error
            
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
            
            Write-CSApiLog "Successfully authenticated. User ID: $userId" -Level Success
            return $true
        } else {
            Write-CSApiLog "Authentication failed: No access_token in response" -Level Error
            Write-CSApiLog "Response content: $($response | ConvertTo-Json -Depth 3)" -Level Error
            
            # Common issues
            Write-CSApiLog "Common issues:" -Level Error
            Write-CSApiLog "  1. Check tenant name is correct (no spaces, exact match, case-sensitive)" -Level Error
            Write-CSApiLog "  2. Verify Client ID and Client Secret are correct (copy exactly from ConnectSecure)" -Level Error
            Write-CSApiLog "  3. Ensure format is: tenant+client_id:client_secret (with + separator)" -Level Error
            Write-CSApiLog "  4. Check Base URL is correct - e.g. https://pod104.myconnectsecure.com" -Level Error
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
    if ([string]::IsNullOrWhiteSpace($name) -and $Company["name.keyword"]) { $name = $Company["name.keyword"] }
    if ([string]::IsNullOrWhiteSpace($name) -and $Company._source["name.keyword"]) { $name = $Company._source["name.keyword"] }
    
    # Fallback: any property containing "name" with a non-empty value
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
                
                # Log first company structure for debugging
                if ($allCompanies.Count -eq 0 -and $first) {
                    $sample = $first | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
                    $joinDelim = ', '
                    Write-CSApiLog ('First company properties: ' + ($sample -join $joinDelim)) -Level Info
                    try {
                        $sampleJson = $first | ConvertTo-Json -Depth 5 -Compress
                        Write-CSApiLog ('First company (raw): ' + $sampleJson.Substring(0, [Math]::Min(500, $sampleJson.Length)) + '...') -Level Info
                    } catch { }
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
}
# Max records to fetch (0 = unlimited). Stops after this - ~5 pages at 5k each. Set higher if you need more.
$script:VulnMaxRecords = 25000
# When using asset_wise_vulnerabilities: aggregate to one row per unique vuln with AffectedHosts (reduces 200k→~50k). Ignored for vulnerabilities_details.
$script:VulnAggregateByVulnerability = $true
# API filter - ConnectSecure may ignore this. Set to "" for no API filter.
$script:VulnSeverityFilter = 'severity.keyword:(Critical OR High)'
# Client-side filter: after download, keep only Critical and High (API filter often ignored). Set $false for all severities.
$script:VulnFilterCriticalHighOnly = $true

function Invoke-ConnectSecureVulnerabilityQuery {
    <#
    .SYNOPSIS
    Shared implementation for fetching vulnerability data from report_queries endpoints.
    .PARAMETER VulnType
    One of: application, external, suppressed
    .PARAMETER MaxRecords
    Override script VulnMaxRecords for this call. 0 = use script default.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('application', 'external', 'suppressed')]
        [string]$VulnType,
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = "",
        [string]$Sort = 'severity.keyword:desc',
        [bool]$FetchAll = $false,
        [int]$MaxRecords = -1
    )

    $endpoint = $script:VulnEndpoints[$VulnType]
    $label = switch ($VulnType) {
        'application' { 'vulnerabilities' }
        'external'    { 'external vulnerabilities' }
        'suppressed'  { 'suppressed vulnerabilities' }
    }

    Write-CSApiLog ('Fetching ' + $label + ' (CompanyId: ' + $CompanyId + ', Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info

    $allData = @()
    $currentSkip = $Skip
    $pageSize = $Limit
    $maxPages = 200  # Safety: 200 x 5000 = 1M records max

    do {
        $queryParams = @{
            limit = $pageSize
            skip  = $currentSkip
            sort  = $Sort
        }

        if ($CompanyId -gt 0) { $queryParams.company_id = $CompanyId }
        if (-not [string]::IsNullOrWhiteSpace($Filter)) { $queryParams.filter = $Filter }

        try {
            $response = Invoke-ConnectSecureRequest -Endpoint $endpoint -QueryParameters $queryParams

            if ($response.status -and $response.data) {
                $allData += $response.data
                $pageNum = [Math]::Floor($currentSkip / $pageSize) + 1
                Write-CSApiLog ('Page ' + $pageNum + ' : retrieved ' + $response.data.Count + ' ' + $label + ' (Total: ' + $allData.Count + ')') -Level Info

                $maxRec = if ($MaxRecords -ge 0) { $MaxRecords } else { $script:VulnMaxRecords }
                if ($maxRec -gt 0 -and $allData.Count -ge $maxRec) {
                    Write-CSApiLog ('Reached record limit (' + $maxRec + ').') -Level Warning
                    break
                }
                if ($FetchAll -eq $true -and $response.data.Count -eq $pageSize) {
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
            software_name = if ($r.software_name) { $r.software_name } else { "" }
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
    Invoke-ConnectSecureVulnerabilityQuery -VulnType 'external' -CompanyId $CompanyId -Limit $Limit -Skip $Skip -Filter $Filter -Sort $Sort -FetchAll $doFetchAll
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
    if ($Raw) { return $data }
    if ($script:VulnEndpoints['suppressed'] -match 'vulnerabilities_details') {
        return ConvertFrom-VulnerabilitiesDetailsFormat -Data $data
    }
    return $data
}

function Get-ConnectSecureRemediationPlan {
    param(
        [int]$CompanyId = 0,
        [int]$Limit = 5000,
        [int]$Skip = 0,
        [string]$Filter = ""
    )

    Write-CSApiLog ('Fetching remediation plan (CompanyId: ' + $CompanyId + ', Limit: ' + $Limit + ', Skip: ' + $Skip + ')...') -Level Info

    $queryParams = @{
        limit = $Limit
        skip = $Skip
    }

    if ($CompanyId -gt 0) {
        $queryParams.company_id = $CompanyId
    }

    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $queryParams.filter = $Filter
    }

    try {
        $response = Invoke-ConnectSecureRequest -Endpoint '/r/asset/get_asset_remediation_plan' -QueryParameters $queryParams
        
        if ($response.status -and $response.data) {
            Write-CSApiLog ('Retrieved ' + $response.data.Count + ' remediation plan records') -Level Success
            return $response.data
        } else {
            Write-CSApiLog 'No remediation plan data returned' -Level Warning
            return @()
        }

    } catch {
        Write-CSApiLog ('Error fetching remediation plan: ' + $_.Exception.Message) -Level Error
        throw
    }
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
    do {
        $queryParams = @{ limit = $pageSize; skip = $currentSkip }
        if ($CompanyId -gt 0) { $queryParams.company_id = $CompanyId }
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
            Write-CSApiLog ('Error fetching assets: ' + $_.Exception.Message) -Level Error
            throw
        }
    } while ($doFetchAll)
    Write-CSApiLog ('Total assets retrieved: ' + $allData.Count) -Level Success
    return $allData
}

function Get-ConnectSecureAssetIdToNameMap {
    <#
    .SYNOPSIS
    Builds a hashtable mapping asset ID -> asset name (hostname, name, or id as fallback).
    Call Get-ConnectSecureAssets first and pass the result.
    #>
    param([array]$Assets)
    $map = @{}
    if (-not $Assets -or $Assets.Count -eq 0) { return $map }
    foreach ($a in $Assets) {
        $id = $null
        if ($null -ne $a.id) { $id = $a.id }
        elseif ($null -ne $a.asset_id) { $id = $a.asset_id }
        elseif ($null -ne $a._id) { $id = $a._id }
        if ($null -eq $id) { continue }
        $key = [string]$id
        if ($map.ContainsKey($key)) { continue }
        $name = $null
        if ($null -ne $a.hostname -and -not [string]::IsNullOrWhiteSpace([string]$a.hostname)) { $name = [string]$a.hostname }
        elseif ($null -ne $a.host_name -and -not [string]::IsNullOrWhiteSpace([string]$a.host_name)) { $name = [string]$a.host_name }
        elseif ($null -ne $a.name -and -not [string]::IsNullOrWhiteSpace([string]$a.name)) { $name = [string]$a.name }
        elseif ($null -ne $a.asset_name -and -not [string]::IsNullOrWhiteSpace([string]$a.asset_name)) { $name = [string]$a.asset_name }
        elseif ($null -ne $a.fqdn_name -and -not [string]::IsNullOrWhiteSpace([string]$a.fqdn_name)) { $name = [string]$a.fqdn_name }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $key }
        $map[$key] = $name
    }
    return $map
}

function Invoke-AssetNameResolution {
    <#
    .SYNOPSIS
    Fetches assets, builds id->name map, and enriches data with asset_names. If asset fetch fails, returns data unchanged.
    #>
    param(
        [array]$Data,
        [int]$CompanyId = 0
    )
    if (-not $Data -or $Data.Count -eq 0) { return $Data }
    try {
        $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
        $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        if ($assetMap.Count -gt 0) {
            return Resolve-AssetIdsToNames -Data $Data -AssetMap $assetMap
        }
    } catch {
        Write-CSApiLog ('Asset name resolution skipped: ' + $_.Exception.Message) -Level Warning
    }
    return $Data
}

function Resolve-AssetIdsToNames {
    <#
    .SYNOPSIS
    Enriches vulnerability data: adds asset_names column by resolving asset_ids using the asset lookup map.
    If asset_ids is present, creates asset_names (semicolon-separated names).
    #>
    param(
        [array]$Data,
        [hashtable]$AssetMap,
        [switch]$ReplaceAssetIds
    )
    if (-not $Data -or $Data.Count -eq 0 -or -not $AssetMap) { return $Data }
    $result = @()
    foreach ($row in $Data) {
        $obj = if ($null -ne $row -and $row.PSObject.Properties['_source']) { $row._source } else { $row }
        $aidVal = $null
        if ($null -ne $obj) {
            if ($obj.PSObject.Properties['asset_ids']) { $aidVal = $obj.asset_ids }
            elseif ($obj -is [hashtable] -and $obj.ContainsKey('asset_ids')) { $aidVal = $obj['asset_ids'] }
        }
        $nameStr = ''
        if ($null -ne $aidVal) {
            $ids = if ($aidVal -is [string]) { ($aidVal -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($aidVal) | Where-Object { $null -ne $_ } }
            $names = @()
            foreach ($id in $ids) {
                $k = [string]$id
                if ($AssetMap.ContainsKey($k)) { $names += $AssetMap[$k] } else { $names += $k }
            }
            $nameStr = $names -join '; '
        }
        # Build new object with asset_names (and optionally replace asset_ids)
        $props = @{}
        if ($obj -is [PSCustomObject]) {
            $obj.PSObject.Properties | ForEach-Object { $props[$_.Name] = $_.Value }
        } elseif ($obj -is [hashtable]) {
            $obj.Keys | ForEach-Object { $props[$_] = $obj[$_] }
        } else {
            $result += $row
            continue
        }
        if ($nameStr) {
            $props['asset_names'] = $nameStr
            if ($ReplaceAssetIds) { $props['asset_ids'] = $nameStr }
        }
        $result += [PSCustomObject]$props
    }
    return $result
}

function Convert-ConnectSecureDataToExcelFormat {
    param(
        [array]$ConnectSecureData,
        [scriptblock]$OnProgress = $null
    )

    Write-CSApiLog 'Converting ConnectSecure data to Excel format...' -Level Info

    $excelData = @()
    $total = if ($ConnectSecureData) { $ConnectSecureData.Count } else { 0 }
    $progressInterval = if ($total -gt 0) { [Math]::Max(100, [Math]::Floor($total / 20)) } else { 99999 }
    $schemaSaved = $false  # Save first record to debug file when fields are missing

    # ConnectSecure native report uses: Host Name, FQDN Name, IP, Software Name, CVE, Problem Name, Solution, Severity
    $hostPaths = @('Host Name','host_name','hostName','hostname','fqdn_name','fqdn','host.hostname','host.host_name','host.name','asset_name','computer_name','device_name','asset.hostname','asset.host_name','machine.hostname','device.hostname','_source.Host Name','_source.host_name','_source.hostname','_source.fqdn_name','_source.host.hostname','_source.asset.hostname')
    $ipPaths = @('IP','ip','ip_address','ipAddress','host.ip','host.ip_address','asset.ip','asset.ip_address','_source.IP','_source.ip','_source.ip_address','_source.host.ip','_source.asset.ip')
    $productPaths = @('Software Name','software_name','softwareName','product','name','cpe','vulnerability.product','vulnerability.software_name','asset.product','_source.Software Name','_source.product','_source.software_name','_source.name')
    $sevPaths = @('Severity','severity','severity.keyword','vulnerability.severity','vulnerability.severity.keyword','_source.Severity','_source.severity','_source.severity.keyword')
    $idx = 0
    foreach ($item in $ConnectSecureData) {
        if ($OnProgress -and $total -gt 0 -and $progressInterval -gt 0 -and ($idx % $progressInterval -eq 0)) {
            $pct = [Math]::Min(99, [Math]::Floor(100 * $idx / $total))
            $msg = 'Converting to Excel format... {0} of {1} ({2} percent)' -f $idx, $total, $pct
            & $OnProgress $msg
        }
        $idx++
        $src = if ($item._source) { $item._source } else { $item }
        # Aggregated format: one row per vuln with AffectedHosts, HostCount
        if ($src._aggregated) {
            $hostVal = $src.AffectedHosts
            if ($src.HostCount) { $hostVal = ('(' + $src.HostCount + ' hosts) ') + $hostVal }
            $ipVal = ""
            $prodVal = $src.software_name
            $sevVal = $src.severity
            $cveVal = $src.problem_name
            $descVal = $src.description
            $epssVal = $src.epss_score
            $fixVal = ""; $evPathVal = ""; $evVerVal = ""
        } else {
        # Path-based lookup first, then deep recursive search as fallback
        $hostVal = Get-ConnectSecureVulnField -Item $src -Paths $hostPaths
        if ($null -eq $hostVal) { $hostVal = Get-ConnectSecureVulnField -Item $item -Paths $hostPaths }
        if ($null -eq $hostVal) { $hostVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Host Name','host_name','hostName','hostname','fqdn_name','asset_name','computer_name','device_name') }
        $ipVal = Get-ConnectSecureVulnField -Item $src -Paths $ipPaths
        if ($null -eq $ipVal) { $ipVal = Get-ConnectSecureVulnField -Item $item -Paths $ipPaths }
        if ($null -eq $ipVal) { $ipVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('ip','ip_address','ipAddress') }
        $prodVal = Get-ConnectSecureVulnField -Item $src -Paths $productPaths
        if ($null -eq $prodVal) { $prodVal = Get-ConnectSecureVulnField -Item $item -Paths $productPaths }
        $sevVal = Get-ConnectSecureVulnField -Item $src -Paths $sevPaths
        if ($null -eq $sevVal) { $sevVal = Get-ConnectSecureVulnField -Item $item -Paths $sevPaths }
        if ($null -eq $sevVal) { $sevVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Severity','severity') }
        $epssVal = Get-ConnectSecureVulnField -Item $src -Paths @('epss_score','epssScore','_source.epss_score')
        if ($null -eq $epssVal) { $epssVal = Get-ConnectSecureVulnField -Item $item -Paths @('epss_score','epssScore') }
        if ($null -eq $epssVal) { $epssVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('epss_score','epssScore') }
        $fixVal = Get-ConnectSecureVulnField -Item $src -Paths @('Solution','solution','fix','remediation','_source.Solution','_source.solution','_source.fix')
        if ($null -eq $fixVal) { $fixVal = Get-ConnectSecureVulnField -Item $item -Paths @('Solution','solution','fix') }
        if ($null -eq $fixVal) { $fixVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Solution','solution','fix','remediation') }
        $evPathVal = Get-ConnectSecureVulnField -Item $src -Paths @('evidence_path','evidencePath','_source.evidence_path')
        if ($null -eq $evPathVal) { $evPathVal = Get-ConnectSecureVulnField -Item $item -Paths @('evidence_path') }
        if ($null -eq $evPathVal) { $evPathVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('evidence_path','evidencePath') }
        $evVerVal = Get-ConnectSecureVulnField -Item $src -Paths @('evidence_version','evidenceVersion','_source.evidence_version')
        if ($null -eq $evVerVal) { $evVerVal = Get-ConnectSecureVulnField -Item $item -Paths @('evidence_version') }
        if ($null -eq $evVerVal) { $evVerVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('evidence_version','evidenceVersion') }
        $cveVal = Get-ConnectSecureVulnField -Item $src -Paths @('CVE','cve','cve_id','_source.CVE','_source.cve')
        if ($null -eq $cveVal) { $cveVal = Get-ConnectSecureVulnField -Item $item -Paths @('CVE','cve','cve_id') }
        if ($null -eq $cveVal) { $cveVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('CVE','cve','cve_id') }
        $descVal = Get-ConnectSecureVulnField -Item $src -Paths @('Problem Name','problem_name','description','summary','_source.Problem Name','_source.problem_name','_source.description')
        if ($null -eq $descVal) { $descVal = Get-ConnectSecureVulnField -Item $item -Paths @('Problem Name','problem_name','description') }
        if ($null -eq $descVal) { $descVal = Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('Problem Name','problem_name','description','summary') }
        }
        $emptyStr = [string]::Empty
        $excelRow = [PSCustomObject]@{
            'Host Name' = if ($hostVal) { $hostVal } else { $emptyStr }
            'IP' = if ($ipVal) { $ipVal } else { $emptyStr }
            'Product' = if ($prodVal) { $prodVal } else { $emptyStr }
            'Severity' = if ($sevVal) { $sevVal } else { $emptyStr }
            'EPSS Score' = if ($null -ne $epssVal) { [double]$epssVal } else { 0.0 }
            'Fix' = if ($fixVal) { $fixVal } else { $emptyStr }
            'Evidence Path' = if ($evPathVal) { $evPathVal } else { $emptyStr }
            'Evidence Version' = if ($evVerVal) { $evVerVal } else { $emptyStr }
            'CVE' = if ($cveVal) { $cveVal } else { $emptyStr }
            'Description' = if ($descVal) { $descVal } else { $emptyStr }
        }
        
        $excelData += $excelRow

        # Auto-capture API schema when Product found but other fields missing (helps diagnose field mapping)
        if (-not $schemaSaved -and $total -gt 0 -and $prodVal -and (-not $hostVal -or -not $ipVal -or -not $sevVal)) {
            $schemaSaved = $true
            try {
                $debugDir = Join-Path $env:LOCALAPPDATA 'VScanMagic'
                if (-not (Test-Path $debugDir)) { New-Item -ItemType Directory -Path $debugDir -Force | Out-Null }
                $debugPath = Join-Path $debugDir 'api-schema-debug.json'
                $first = $ConnectSecureData[0]
                $first | ConvertTo-Json -Depth 10 | Set-Content -Path $debugPath -Encoding UTF8
                $script:__schemaPaths = @()
                function Walk-Object { param($o,$prefix)
                    if (-not $o) { return }
                    $o.PSObject.Properties | ForEach-Object {
                        $k = $_.Name; $v = $_.Value
                        $path = if ($prefix) { $prefix + '.' + $k } else { $k }
                        if ($v -is [string] -or $v -is [int] -or $v -is [double] -or $v -is [bool]) { $script:__schemaPaths += $path }
                        elseif ($v -is [PSCustomObject] -and ($null -eq $prefix -or $prefix.Length -lt 25)) { Walk-Object $v $path }
                    }
                }
                Walk-Object $first $null
                $keyList = $script:__schemaPaths -join ', '
                Write-CSApiLog ('Schema debug: Saved to ' + $debugPath + ' | Paths: ' + $keyList) -Level Warning
            } catch { Write-CSApiLog ('Could not save schema debug: ' + $_.Message) -Level Warning }
        }
    }

    Write-CSApiLog ('Converted ' + $excelData.Count + ' records to Excel format') -Level Success
    return $excelData
}

function Export-ConnectSecureDataToExcel {
    param(
        [array]$Data,
        [string]$OutputPath,
        [string]$SheetName = 'Vulnerabilities',
        [scriptblock]$OnProgress = $null
    )

    Write-CSApiLog ('Exporting data to Excel: ' + $OutputPath) -Level Info

    $excel = $null
    $workbook = $null

    try {
        if ($OnProgress) { & $OnProgress 'Opening Excel...' }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = $SheetName

        # Derive headers from API data (union of all rows - API output is source of truth)
        $headers = @()
        if ($Data -and $Data.Count -gt 0) {
            $allKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($item in $Data) {
                $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
                if ($null -ne $obj) {
                    if ($obj -is [hashtable]) {
                        foreach ($k in $obj.Keys) { $null = $allKeys.Add([string]$k) }
                    } else {
                        $obj.PSObject.Properties.Name | Where-Object { $_ -ne '_source' } | ForEach-Object { $null = $allKeys.Add($_) }
                    }
                }
            }
            $headers = [string[]]$allKeys
        }

        for ($col = 1; $col -le $headers.Count; $col++) {
            $r = $worksheet.Cells.Item(1, $col)
            $r.Value2 = [string]$headers[$col - 1]
            $r.Font.Bold = $true
        }

        # Write data row by row - API output as source of truth
        $row = 2
        $total = if ($Data) { $Data.Count } else { 0 }
        $progressInterval = if ($total -gt 0) { [Math]::Max(100, [Math]::Floor($total / 20)) } else { 99999 }
        $written = 0
        foreach ($item in $Data) {
            $obj = if ($null -ne $item -and $item.PSObject.Properties['_source']) { $item._source } else { $item }
            for ($col = 1; $col -le $headers.Count; $col++) {
                $key = $headers[$col - 1]
                $val = $null
                if ($null -ne $obj) {
                    if ($obj -is [hashtable] -and $obj.ContainsKey($key)) { $val = $obj[$key] }
                    elseif ($obj.PSObject.Properties[$key]) { $val = $obj.PSObject.Properties[$key].Value }
                }
                if ($null -eq $val) { $val = '' }
                elseif ($val -is [Array] -or ($val -is [System.Collections.IEnumerable] -and $val -isnot [string])) {
                    $arr = @($val) | Where-Object { $null -ne $_ }
                    $val = ($arr | ForEach-Object { $_.ToString() }) -join '; '
                }
                elseif ($val -is [PSCustomObject] -or $val -is [hashtable]) {
                    try { $val = ($val | ConvertTo-Json -Compress -Depth 2) } catch { $val = $val.ToString() }
                }
                elseif ($val -isnot [string]) { $val = $val.ToString() }
                $cellVal = [string]$val
                $targetCell = $worksheet.Cells.Item([int]$row, [int]$col)
                $targetCell.Value2 = $cellVal
            }
            $row++
            $written++
            if ($OnProgress -and $total -gt 0 -and $progressInterval -gt 0 -and ($written % $progressInterval -eq 0)) {
                $pct = [Math]::Min(99, [Math]::Floor(100 * $written / $total))
                $msg = [string]('Exporting to Excel... ' + $written + ' of ' + $total + ' rows (' + $pct + ' pct)')
                & $OnProgress $msg
            }
        }

        # Auto-fit columns
        if ($OnProgress) { & $OnProgress 'Formatting columns...' }
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null

        # Save workbook
        if ($OnProgress) { & $OnProgress 'Saving Excel file...' }
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $workbook.SaveAs($OutputPath)
        $workbook.Close($false)

        Write-CSApiLog ('Excel file saved: ' + $OutputPath) -Level Success

    } catch {
        Write-CSApiLog ('Error exporting to Excel: ' + $_.Exception.Message) -Level Error
        throw
    } finally {
        if ($workbook) {
            try {
                $workbook.Close($false)
                Clear-ComObject $workbook
            } catch { }
        }
        if ($excel) {
            try {
                $excel.Quit()
                Clear-ComObject $excel
            } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Clear-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {
            # Ignore errors
        }
    }
}

# --- Report Generation from ConnectSecure Data (no API server needed) ---

function Get-ConnectSecureVulnField {
    param([object]$Item, [string[]]$Paths)

    # Helper: extract value from object (handle .keyword sub-property, arrays like [ios])
    $extractVal = {
        param($o)
        if ($null -eq $o -or $o -eq '') { return $null }
        if ($o -is [string] -or $o -is [int] -or $o -is [long] -or $o -is [double] -or $o -is [bool]) { return $o }
        if ($o -is [Array]) {
            $first = $o | Where-Object { $null -ne $_ -and $_ -ne '' } | Select-Object -First 1
            if ($first) {
                if ($first -is [string] -or $first -is [int] -or $first -is [long] -or $first -is [double]) { return $first }
                try { if ($first.PSObject.Properties['keyword']) { return $first.keyword } } catch { }
                return $first.ToString()
            }
            return $null
        }
        try { if ($o.PSObject.Properties['keyword']) { return $o.keyword } } catch { }
        return $o.ToString()
    }

    foreach ($p in $Paths) {
        try {
            $o = $Item
            # Try literal property name first (e.g. severity.keyword as single prop)
            if ($p.Contains('.')) {
                $litVal = $null
                try { $litVal = $o.PSObject.Properties[$p].Value } catch { }
                if ($null -ne $litVal -and $litVal -ne '') {
                    $v = & $extractVal $litVal
                    if ($null -ne $v) { return $v }
                }
            }
            # Traverse path
            $parts = $p -split '\.'
            foreach ($part in $parts) {
                if ($null -eq $o) { break }
                try { $o = $o.$part } catch { $o = $null; break }
            }
            if ($null -ne $o -and $o -ne '') {
                $v = & $extractVal $o
                if ($null -ne $v) { return $v }
            }
        } catch { }
    }
    return $null
}

function Get-ConnectSecureVulnFieldDeep {
    <# Recursively search object tree for first match of any known field name. Use when path-based lookup fails. #>
    param([object]$Item, [string[]]$FieldNames, [int]$MaxDepth = 5)
    if ($null -eq $Item -or $MaxDepth -le 0) { return $null }
    try {
        foreach ($prop in $Item.PSObject.Properties) {
            $name = $prop.Name
            $val = $prop.Value
            if ($name -in $FieldNames -and $null -ne $val -and $val -ne '') {
                if ($val -is [string] -or $val -is [int] -or $val -is [long] -or $val -is [double]) { return $val }
                if ($val -is [Array]) {
                    $first = $val | Where-Object { $null -ne $_ -and $_ -ne '' } | Select-Object -First 1
                    if ($first -is [string]) { return $first }
                    if ($null -ne $first) { return $first.ToString() }
                    return $null
                }
                try { if ($val.PSObject.Properties['keyword']) { return $val.keyword } } catch { }
                return $val.ToString()
            }
        }
        foreach ($prop in $Item.PSObject.Properties) {
            $val = $prop.Value
            if ($null -ne $val -and $val -is [System.Management.Automation.PSCustomObject]) {
                $found = Get-ConnectSecureVulnFieldDeep -Item $val -FieldNames $FieldNames -MaxDepth ($MaxDepth - 1)
                if ($null -ne $found) { return $found }
            }
        }
    } catch { }
    return $null
}

function Convert-ConnectSecureToVulnData {
    param([array]$ConnectSecureData)

    # Align with ConnectSecure native report: Host Name, FQDN Name, IP, Software Name, CVE, Problem Name, Solution, Severity
    $hostPaths = @('Host Name','host_name','hostName','hostname','fqdn_name','fqdn','host.hostname','host.host_name','asset_name','computer_name','device_name','asset.hostname','asset.host_name','machine.hostname','host.hostname','device.hostname','_source.Host Name','_source.host_name','_source.hostname','_source.fqdn_name','_source.asset.hostname')
    $ipPaths = @('IP','ip','ip_address','ipAddress','host.ip','asset.ip','asset.ip_address','machine.ip','device.ip','_source.IP','_source.ip','_source.ip_address','_source.asset.ip')
    $productPaths = @('Software Name','software_name','softwareName','product','name','cpe','vulnerability.product','vulnerability.software_name','asset.product','_source.Software Name','_source.product','_source.software_name','_source.name')
    $sevPaths = @('Severity','severity','severity.keyword','vulnerability.severity','vulnerability.severity.keyword','_source.Severity','_source.severity','_source.severity.keyword')
    $epssPaths = @('epss_score','epssScore','vulnerability.epss_score','_source.epss_score')

    $vulnData = @()
    $debugLogged = $false
    foreach ($item in $ConnectSecureData) {
        $src = if ($item._source) { $item._source } else { $item }
        if ($src._aggregated) {
            $hostName = $src.AffectedHosts
            if ($src.HostCount) { $hostName = ('(' + $src.HostCount + ' hosts) ') + $hostName }
            $ip = ""
            $product = $src.software_name
            $severity = $src.severity
            $epssScore = if ($null -ne $src.epss_score) { [double]$src.epss_score } else { 0.0 }
            $critical = if ($severity -eq 'Critical') { 1 } else { 0 }
            $high = if ($severity -eq 'High') { 1 } else { 0 }
            $medium = if ($severity -eq 'Medium') { 1 } else { 0 }
            $low = if ($severity -eq 'Low') { 1 } else { 0 }
            $vulnData += [PSCustomObject]@{ 'Host Name' = $hostName; 'IP' = $ip; 'Username' = ''; 'Product' = $product; 'Critical' = $critical; 'High' = $high; 'Medium' = $medium; 'Low' = $low; 'Vulnerability Count' = [Math]::Max(1, $src.HostCount); 'EPSS Score' = $epssScore }
            continue
        }
        $hostName = (Get-ConnectSecureVulnField -Item $src -Paths $hostPaths)
        if ($null -eq $hostName) { $hostName = (Get-ConnectSecureVulnField -Item $item -Paths $hostPaths) }
        if ($null -eq $hostName) { $hostName = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('host_name','hostName','hostname','asset_name','computer_name','device_name')) }
        if ($null -eq $hostName) { $hostName = "" }
        $ip = (Get-ConnectSecureVulnField -Item $src -Paths $ipPaths)
        if ($null -eq $ip) { $ip = (Get-ConnectSecureVulnField -Item $item -Paths $ipPaths) }
        if ($null -eq $ip) { $ip = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('ip','ip_address','ipAddress')) }
        if ($null -eq $ip) { $ip = "" }
        $product = (Get-ConnectSecureVulnField -Item $src -Paths $productPaths)
        if ($null -eq $product) { $product = (Get-ConnectSecureVulnField -Item $item -Paths $productPaths) }
        if ($null -eq $product) { $product = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('product','software_name','softwareName','name')) }
        if ($null -eq $product) { $product = "" }
        $sevRaw = (Get-ConnectSecureVulnField -Item $src -Paths $sevPaths)
        if ($null -eq $sevRaw) { $sevRaw = (Get-ConnectSecureVulnField -Item $item -Paths $sevPaths) }
        if ($null -eq $sevRaw) { $sevRaw = (Get-ConnectSecureVulnFieldDeep -Item $item -FieldNames @('severity')) }
        $severity = if ($sevRaw) { if ($sevRaw -is [string]) { $sevRaw } else { $sevRaw.keyword } } else { 'Medium' }
        $epssRaw = (Get-ConnectSecureVulnField -Item $src -Paths $epssPaths)
        if ($null -eq $epssRaw) { $epssRaw = (Get-ConnectSecureVulnField -Item $item -Paths $epssPaths) }
        $epssScore = if ($null -ne $epssRaw) { [double]$epssRaw } else { 0.0 }
        if (-not $debugLogged -and $ConnectSecureData.Count -gt 0 -and ($hostName -eq '' -or $ip -eq '') -and $product -ne '') {
            $first = $ConnectSecureData[0]
            $keys = @()
            if ($first._source) { $first._source.PSObject.Properties | ForEach-Object { $keys += '_source.' + $_.Name } }
            $first.PSObject.Properties | Where-Object { $_.Name -ne '_source' } | ForEach-Object { $keys += $_.Name }
            Write-CSApiLog ('ConnectSecure vuln record sample keys (host/ip empty): ' + ($keys -join ', ')) -Level Warning
            $debugLogged = $true
        }
        $critical = if ($severity -eq 'Critical') { 1 } else { 0 }
        $high = if ($severity -eq 'High') { 1 } else { 0 }
        $medium = if ($severity -eq 'Medium') { 1 } else { 0 }
        $low = if ($severity -eq 'Low') { 1 } else { 0 }
        $vulnData += [PSCustomObject]@{
            'Host Name' = $hostName
            'IP' = $ip
            'Username' = ''
            'Product' = $product
            'Critical' = $critical
            'High' = $high
            'Medium' = $medium
            'Low' = $low
            'Vulnerability Count' = 1
            'EPSS Score' = $epssScore
        }
    }
    return $vulnData
}

function New-PendingEPSSReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [array]$VulnerabilityData = $null, [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $allVulns = if ($null -ne $VulnerabilityData) { $VulnerabilityData } else { Get-ConnectSecureVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -Raw }
    if ($null -eq $allVulns) { $allVulns = @() }
    $pendingEPSS = $allVulns | Where-Object { $_.epss_score -and [double]$_.epss_score -gt 0 }
    if ($null -eq $pendingEPSS) { $pendingEPSS = @() }
    $pendingEPSS = Invoke-AssetNameResolution -Data $pendingEPSS -CompanyId $CompanyId
    Export-ConnectSecureDataToExcel -Data $pendingEPSS -OutputPath $OutputPath -SheetName 'Pending Remediation EPSS' -OnProgress $null
}

function New-AllVulnerabilitiesReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [array]$VulnerabilityData = $null, [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $allVulns = if ($null -ne $VulnerabilityData) { $VulnerabilityData } else { Get-ConnectSecureVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -Raw }
    if ($null -eq $allVulns) { $allVulns = @() }
    $allVulns = Invoke-AssetNameResolution -Data $allVulns -CompanyId $CompanyId
    Export-ConnectSecureDataToExcel -Data $allVulns -OutputPath $OutputPath -SheetName 'All Vulnerabilities' -OnProgress $null
}

function New-ExternalVulnerabilitiesReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $externalVulns = Get-ConnectSecureExternalVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll
    if ($null -eq $externalVulns) { $externalVulns = @() }
    $externalVulns = Invoke-AssetNameResolution -Data $externalVulns -CompanyId $CompanyId
    Export-ConnectSecureDataToExcel -Data $externalVulns -OutputPath $OutputPath -SheetName 'External Vulnerabilities' -OnProgress $null
}

function New-SuppressedVulnerabilitiesReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $suppressedVulns = Get-ConnectSecureSuppressedVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -Raw
    if ($null -eq $suppressedVulns) { $suppressedVulns = @() }
    $suppressedVulns = Invoke-AssetNameResolution -Data $suppressedVulns -CompanyId $CompanyId
    Export-ConnectSecureDataToExcel -Data $suppressedVulns -OutputPath $OutputPath -SheetName 'Suppressed Vulnerabilities' -OnProgress $null
}

function New-ExecutiveSummaryReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [string]$ScanDate, [array]$VulnerabilityData = $null, [int]$TopCount = 10, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $allVulns = if ($null -ne $VulnerabilityData) { $VulnerabilityData } else { Get-ConnectSecureVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll }
    if ($null -eq $allVulns) { $allVulns = @() }
    $vulnData = Convert-ConnectSecureToVulnData -ConnectSecureData $allVulns
    $top10 = Get-Top10Vulnerabilities -VulnData $vulnData -Count $TopCount
    New-WordReport -OutputPath $OutputPath -ClientName $ClientName -ScanDate $ScanDate -Top10Data $top10 `
        -TimeEstimates $null -IsRMITPlus $false -GeneralRecommendations @() -ReportTitle 'Executive Summary'
}

# --- Report Builder Functions (ConnectSecure native report generation) ---
# Create job on CS, poll until ready, download XLSX/DOCX from CS.

function Get-ConnectSecureStandardReports {
    <#
    .SYNOPSIS
    Gets list of available standard reports from ConnectSecure report builder.
    Swagger: GET /report_builder/standard_reports (requires isGlobal query param).
    Response: message[].Reports[].reports[] with id, reportType, templateFile.
    #>
    param([bool]$IsGlobal = $false)
    $queryParams = @{ isGlobal = $IsGlobal.ToString().ToLower() }
    try {
        $response = Invoke-ConnectSecureRequest -Endpoint '/report_builder/standard_reports' -Method GET -QueryParameters $queryParams
        if (-not $response.status) { return @() }
        $reports = @()
        if ($response.message) {
            $msg = $response.message
            if ($msg -is [Array]) {
                foreach ($m in $msg) {
                    if ($m.Reports) {
                        foreach ($r in $m.Reports) {
                            if ($r.reports) { $reports += $r.reports }
                        }
                    }
                }
            }
        }
        if ($response.data) { $reports += $response.data }
        Write-CSApiLog ('Standard reports: ' + $reports.Count + ' items') -Level Success
        return $reports
    } catch {
        Write-CSApiLog ('standard_reports: ' + $_.Exception.Message) -Level Warning
        return @()
    }
}

function New-ConnectSecureReportJob {
    <#
    .SYNOPSIS
    Creates a report generation job in ConnectSecure.
    Swagger: POST /report_builder/create_report_job
    .PARAMETER ReportType
    Report id (string) or reportType from standard_reports, e.g. pending_remediation_epss or id from API.
    .PARAMETER CompanyId
    Company ID (0 for all companies).
    .PARAMETER ReportFormat
    Format: xlsx, docx, pdf.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$ReportType,
        [int]$CompanyId = 0,
        [string]$ReportFormat = 'xlsx',
        [hashtable]$AdditionalParams = @{}
    )

    $body = @{ company_id = $CompanyId; report_format = $ReportFormat } + $AdditionalParams
    if ($ReportType -is [int]) {
        $body.report_id = $ReportType
    } elseif ($ReportType -match '^\d+$') {
        $body.report_id = $ReportType
    } elseif ($ReportType -match '^[a-fA-F0-9]{32}$') {
        # Standard report ID (hex) from report_builder/standard_reports
        $body.report_id = $ReportType.ToString()
    } else {
        $body.reportType = $ReportType.ToString()
    }

    try {
        $response = Invoke-ConnectSecureRequest -Endpoint '/report_builder/create_report_job' -Method POST -Body $body
        if (-not $response.status) { throw 'create_report_job returned status=false' }

        # Extract job_id from various possible response shapes (API structure varies by ConnectSecure version)
        $jobId = $null
        try {
            if ($null -ne $response.data) {
                $d = $response.data
                if ($null -ne $d.job_id) { $jobId = $d.job_id }
                elseif ($null -ne $d.id) { $jobId = $d.id }
                elseif ($null -ne $d.jobId) { $jobId = $d.jobId }
                elseif ($d -is [string]) { $jobId = $d }
                elseif ($d -is [int] -or $d -is [long]) { $jobId = $d.ToString() }
            }
            if ([string]::IsNullOrWhiteSpace($jobId) -and $null -ne $response.job_id) { $jobId = $response.job_id }
            if ([string]::IsNullOrWhiteSpace($jobId) -and $null -ne $response.id) { $jobId = $response.id }
            if ([string]::IsNullOrWhiteSpace($jobId) -and $null -ne $response.message) {
                $msg = $response.message
                if ($msg -is [string] -and $msg -match '^\d+$') { $jobId = $msg }
                elseif ($null -ne $msg.job_id) { $jobId = $msg.job_id }
            }
        } catch { /* ignore property access errors */ }

        if (-not [string]::IsNullOrWhiteSpace($jobId)) {
            Write-CSApiLog ('Report job created: ' + $jobId) -Level Success
            return $jobId.ToString()
        }

        # Debug: log actual response structure when job_id not found (helps diagnose API differences)
        try {
            $respJson = $response | ConvertTo-Json -Depth 6 -Compress
            Write-CSApiLog ('create_report_job response (no job_id found): ' + $respJson) -Level Warning
        } catch { }
        throw 'No job_id in response'
    } catch {
        Write-CSApiLog ('create_report_job: ' + $_.Exception.Message) -Level Error
        throw 'ConnectSecure create_report_job failed. Ensure report type (id or reportType from standard_reports) is valid.'
    }
}

function Get-ConnectSecureReportJobStatus {
    <#
    .SYNOPSIS
    Swagger: GET /r/company/report_jobs_view/{id} - returns job_id, type, status, description, company_id, etc.
    #>
    param([Parameter(Mandatory=$true)][string]$JobId)
    $response = Invoke-ConnectSecureRequest -Endpoint ('/r/company/report_jobs_view/' + $JobId) -Method GET
    if (-not $response) { throw 'No response from report_jobs_view' }
    return $response
}

function Get-ConnectSecureReportLink {
    <#
    .SYNOPSIS
    Swagger: GET /report_builder/get_report_link - requires job_id and isGlobal (query params).
    Response: status, message (string - may contain download URL).
    For company-scoped reports, company_id may be required.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$JobId,
        [bool]$IsGlobal = $false,
        [int]$CompanyId = 0
    )
    $isGlob = $IsGlobal.ToString().ToLower()
    $paramVariants = @(
        @{ job_id = $JobId; isGlobal = $isGlob },
        @{ jobId = $JobId; isGlobal = $isGlob }
    )
    if (-not $IsGlobal -and $CompanyId -ne 0) {
        foreach ($p in $paramVariants) { $p['company_id'] = $CompanyId }
    }
    $endpoints = @('/report_builder/get_report_link', '/r/report_builder/get_report_link')
    $response = $null
    $lastErr = $null
    foreach ($ep in $endpoints) {
        foreach ($qp in $paramVariants) {
            try {
                $response = Invoke-ConnectSecureRequest -Endpoint $ep -Method GET -QueryParameters $qp -RetryCount 6
                if ($response) { break }
            } catch { $lastErr = $_; continue }
        }
        if ($response) { break }
    }
    if (-not $response) { if ($lastErr) { throw $lastErr }; throw 'get_report_link failed' }
    if (-not $response.status) { throw 'get_report_link returned status=false' }
    $url = $response.message
    if ([string]::IsNullOrWhiteSpace($url)) { $url = $response.data.download_url }
    if ([string]::IsNullOrWhiteSpace($url)) { $url = $response.data.url }
    if ([string]::IsNullOrWhiteSpace($url)) { $url = $response.data.link }
    if ([string]::IsNullOrWhiteSpace($url)) { throw 'No download URL in get_report_link response' }
    return $url
}

function Invoke-ConnectSecureReportsBatch {
    <#
    .SYNOPSIS
    Generates standard reports. Tries ConnectSecure Report Builder first for asset-level data (Host Name, IP, CVE per row).
    Falls back to data APIs (application_vulnerabilities, etc.) when Report Builder is unavailable.
    .PARAMETER Reports
    Array of @{ Type; Name; Ext } - internal report types and display names.
    .PARAMETER OutputPathTemplate
    Scriptblock: param($report) returns full output path for that report.
    .PARAMETER CompanyId
    .PARAMETER ClientName
    .PARAMETER ScanDate
    .PARAMETER OnProgress
    Optional scriptblock: param($message) - called to update progress UI.
    #>
    param(
        [Parameter(Mandatory=$true)][array]$Reports,
        [Parameter(Mandatory=$true)][scriptblock]$OutputPathTemplate,
        [int]$CompanyId = 0,
        [string]$ClientName = 'Client',
        [string]$ScanDate = "",
        [int]$TopCount = 10,
        [scriptblock]$OnProgress = $null,
        [int]$DebugLimit = 0
    )

    function Update-Prog { param($m) if ($OnProgress) { & $OnProgress $m } }

    # Standard reports only - try Report Builder first (asset-level: Host, IP, CVE per row), fall back to data APIs
    $standardTypes = @('all-vulnerabilities', 'suppressed-vulnerabilities', 'external-vulnerabilities', 'executive-summary', 'pending-epss')
    $reportTypes = $Reports | ForEach-Object { $_.Type } | Select-Object -Unique

    foreach ($rt in $reportTypes) {
        if ($rt -notin $standardTypes) {
            throw ('Unknown report type: ' + $rt + '. Standard reports: all-vulnerabilities, suppressed-vulnerabilities, external-vulnerabilities, executive-summary, pending-epss')
        }
    }

    $script:CSReportBuilderUnavailable = $false

    # Shared vuln data: fetch only when Report Builder fails (avoids 100k+ pull when Report Builder succeeds)
    $sharedVulnTypes = @('all-vulnerabilities', 'executive-summary', 'pending-epss')
    $allVulns = $null

    $succeeded = [System.Collections.ArrayList]::new()
    $failed = [System.Collections.ArrayList]::new()

    foreach ($report in $Reports) {
        $path = & $OutputPathTemplate $report
        Update-Prog ('Generating ' + $report.Name + '...')
        try {
            $reportFormat = if ($report.Ext -eq 'docx') { 'docx' } else { 'xlsx' }
            $usedReportBuilder = Invoke-ConnectSecureReportBuilderDownload -InternalReportType $report.Type -OutputPath $path -CompanyId $CompanyId -ReportFormat $reportFormat -OnProgress $OnProgress
            if (-not $usedReportBuilder) {
                # Fetch shared data only when needed (Report Builder failed)
                if ($report.Type -in $sharedVulnTypes) {
                    if ($null -eq $allVulns) {
                        Update-Prog 'Fetching vulnerability data (Report Builder unavailable)...'
                        $fetchLimit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
                        $fetchAll = ($DebugLimit -le 0)
                        $allVulns = Get-ConnectSecureVulnerabilities -CompanyId $CompanyId -Limit $fetchLimit -FetchAll:$fetchAll
                        if ($null -eq $allVulns) { $allVulns = @() }
                        Write-CSApiLog ('Fetched ' + $allVulns.Count + ' vulnerabilities for local fallback') -Level Success
                    }
                Invoke-LocalReportFallback -InternalReportType $report.Type -OutputPath $path -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate -VulnerabilityData $allVulns -TopCount $TopCount -OnProgress $OnProgress -DebugLimit $DebugLimit
            } else {
                Invoke-LocalReportFallback -InternalReportType $report.Type -OutputPath $path -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate -VulnerabilityData $null -TopCount $TopCount -OnProgress $OnProgress -DebugLimit $DebugLimit
            }
            }
            Write-CSApiLog ('Generated: ' + $path) -Level Success
            $null = $succeeded.Add($report)
        } catch {
            $errText = if ($null -eq $_.Exception.Message) { 'Unknown error' } elseif ($_.Exception.Message -is [array]) { ($_.Exception.Message -join '; ') } else { [string]$_.Exception.Message }
            Write-CSApiLog ('Report failed (' + $report.Name + '): ' + $errText) -Level Error
            $null = $failed.Add([PSCustomObject]@{ Report = $report; Path = $path; Error = $errText })
        }
    }

    return @{ Succeeded = [array]$succeeded; Failed = [array]$failed }
}

function Invoke-ConnectSecureReportDownload {
    <#
    .SYNOPSIS
    Creates a report job, polls until ready, and downloads the file. (Single-report; batch flow uses Invoke-ConnectSecureReportsBatch.)
    .PARAMETER ReportType
    ConnectSecure report type (e.g. pending_epss, executive_summary).
    .PARAMETER OutputPath
    Full path to save the downloaded file.
    .PARAMETER CompanyId
    Company ID (0 for all).
    .PARAMETER ReportFormat
    xlsx, docx, or pdf.
    .PARAMETER PollIntervalSeconds
    Seconds between status checks.
    .PARAMETER MaxWaitSeconds
    Maximum time to wait for job completion.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ReportType,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        [int]$CompanyId = 0,
        [string]$ReportFormat = 'xlsx',
        [int]$PollIntervalSeconds = 5,
        [int]$MaxWaitSeconds = 300
    )

    $jobId = New-ConnectSecureReportJob -ReportType $ReportType -CompanyId $CompanyId -ReportFormat $ReportFormat
    $start = Get-Date

    while ($true) {
        $elapsed = ((Get-Date) - $start).TotalSeconds
        if ($elapsed -ge $MaxWaitSeconds) {
            throw ('Report job timed out after ' + $MaxWaitSeconds + ' seconds')
        }

        $status = Get-ConnectSecureReportJobStatus -JobId $jobId
        $jobStatus = $null
        if ($status.data) {
            $jobStatus = $status.data.status
            if ([string]::IsNullOrWhiteSpace($jobStatus)) { $jobStatus = $status.data.job_status }
        }

        if ($jobStatus -eq 'completed' -or $jobStatus -eq 'complete' -or $jobStatus -eq 'ready') {
            break
        }
        if ($jobStatus -eq 'failed' -or $jobStatus -eq 'error') {
            $errMsg = if ($status.data.error_message) { $status.data.error_message } else { 'Report job failed' }
            throw $errMsg
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }

    $downloadUrl = Get-ConnectSecureReportLink -JobId $jobId -IsGlobal ($CompanyId -eq 0)
    if (-not $downloadUrl.StartsWith('http')) {
        $slash = [char]47
        $base = $script:ConnectSecureConfig.BaseUrl.TrimEnd($slash)
        $path = $downloadUrl.TrimStart($slash)
        $downloadUrl = $base + '/' + $path
    }

    # Download file (use same auth headers for ConnectSecure link)
    $bearerToken = 'Bearer ' + $script:ConnectSecureConfig.AccessToken
    $headers = @{
        'Authorization' = $bearerToken
    }
    if (-not [string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.UserId)) {
        $headers['X-USER-ID'] = $script:ConnectSecureConfig.UserId.ToString()
    }

    Invoke-WebRequest -Uri $downloadUrl -Method GET -Headers $headers -OutFile $OutputPath -UseBasicParsing
    Write-CSApiLog ('Downloaded report from ConnectSecure to ' + $OutputPath) -Level Success
}

function Invoke-ConnectSecureReportBuilderDownload {
    <#
    .SYNOPSIS
    Uses ConnectSecure Report Builder to generate a report with asset-level data (Host Name, IP, CVE per row).
    Polls get_report_link (avoids report_jobs_view which often returns 502).
    Returns $true if successful, $false if report builder unavailable (404) or fails - caller should fall back to local generation.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$InternalReportType,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [int]$CompanyId = 0,
        [string]$ReportFormat = 'xlsx',
        [int]$PollIntervalSeconds = 5,
        [int]$MaxWaitSeconds = 420,
        [scriptblock]$OnProgress = $null
    )
    function Update-Prog { param($m) if ($OnProgress) { & $OnProgress $m } }

    if ($script:CSReportBuilderUnavailable) { return $false }
    $reportId = $script:CSReportIdMap[$InternalReportType]
    if (-not $reportId) { return $false }

    $ext = if ($ReportFormat -eq 'docx') { 'docx' } elseif ($ReportFormat -eq 'pdf') { 'pdf' } else { 'xlsx' }
    try {
        Update-Prog 'Requesting report from ConnectSecure Report Builder...'
        $jobId = New-ConnectSecureReportJob -ReportType $reportId -CompanyId $CompanyId -ReportFormat $ext
        $start = Get-Date
        $isGlobal = ($CompanyId -eq 0)

        # Poll get_report_link (404 may mean report not ready yet for large datasets)
        Start-Sleep -Seconds 15
        $downloadUrl = $null
        $retries = 0
        $maxRetries = 30
        while ($true) {
            $elapsed = ((Get-Date) - $start).TotalSeconds
            if ($elapsed -ge $MaxWaitSeconds) {
                Write-CSApiLog ('Report job timed out after ' + $MaxWaitSeconds + ' seconds') -Level Warning
                return $false
            }
            Update-Prog ('Waiting for report... (' + [int]$elapsed + 's)')
            try {
                $downloadUrl = Get-ConnectSecureReportLink -JobId $jobId -IsGlobal $isGlobal -CompanyId $CompanyId
                if (-not [string]::IsNullOrWhiteSpace($downloadUrl)) { break }
            } catch {
                $retries++
                if ($retries -ge $maxRetries) { throw }
                $delay = [Math]::Min(5 + ($retries * 3), 30)
                Write-CSApiLog ('Report not ready, retrying in ' + $delay + 's (' + $retries + '/' + $maxRetries + ')...') -Level Warning
                Start-Sleep -Seconds $delay
            }
            Start-Sleep -Seconds $PollIntervalSeconds
        }
        if (-not $downloadUrl.StartsWith('http')) {
            $slash = [char]47
            $base = $script:ConnectSecureConfig.BaseUrl.TrimEnd($slash)
            $path = $downloadUrl.TrimStart($slash)
            $downloadUrl = $base + '/' + $path
        }
        Update-Prog 'Downloading report...'
        $bearerToken = 'Bearer ' + $script:ConnectSecureConfig.AccessToken
        $headers = @{ 'Authorization' = $bearerToken }
        if (-not [string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.UserId)) {
            $headers['X-USER-ID'] = $script:ConnectSecureConfig.UserId.ToString()
        }
        Invoke-WebRequest -Uri $downloadUrl -Method GET -Headers $headers -OutFile $OutputPath -UseBasicParsing
        Write-CSApiLog ('Report Builder: Downloaded asset-level report to ' + $OutputPath) -Level Success
        return $true
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '404|Not Found|report builder') { $script:CSReportBuilderUnavailable = $true }
        Write-CSApiLog ('Report Builder failed (' + $InternalReportType + '): ' + $msg) -Level Warning
        return $false
    }
}

# When CS report builder returns 404, skip attempts for remaining reports (reduces failed requests)
$script:CSReportBuilderUnavailable = $false

# Standard report IDs from report_builder/standard_reports - used for asset-level data (Host Name, IP, CVE per row)
$script:CSReportIdMap = @{
    'all-vulnerabilities' = 'f836d6a4e4d54ac6a9d2967254796373'
    'suppressed-vulnerabilities' = '1d091564830b44c485a0ddc35ace9ac6'
    'external-vulnerabilities' = '01beb6b930744e11b690bb9dc25118fb'
    'executive-summary' = '1cd4f45884264d15bee4173dc58b6a57'
    'pending-epss' = '85d4913c0dbc4fc782b858f0d27dd180'
}

# Report type mapping: our internal type -> possible ConnectSecure report_type/report_id values to try
# ConnectSecure standard reports: All Vulnerabilities, Suppressed Vulnerabilities, External Scan, Executive Summary Report, Pending Remediation EPSS Score Reports
$script:CSReportTypeMap = @{
    'pending-epss' = @('pending_remediation_epss_score_reports','pending_remediation_epss','pending_epss','Pending Remediation EPSS Score Reports')
    'executive-summary' = @('executive_summary_report','executive_summary','Executive Summary Report')
    'all-vulnerabilities' = @('all_vulnerabilities_report','all_vulnerabilities','All Vulnerabilities Report')
    'external-vulnerabilities' = @('external_scan','external_vulnerabilities','External Scan')
    'suppressed-vulnerabilities' = @('suppressed_vulnerabilities','Suppressed Vulnerabilities')
}

function Invoke-ConnectSecureReportDownloadOrFallback {
    <#
    .SYNOPSIS
    Generates a standard report from ConnectSecure data APIs (no report builder).
    .PARAMETER InternalReportType
    pending-epss, executive-summary, all-vulnerabilities, external-vulnerabilities, suppressed-vulnerabilities
    #>
    param(
        [Parameter(Mandatory=$true)][string]$InternalReportType,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [int]$CompanyId = 0,
        [string]$ClientName = 'Client',
        [string]$ScanDate = ''
    )
    Invoke-LocalReportFallback -InternalReportType $InternalReportType -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate
}

function Invoke-LocalReportFallback {
    <#
    .SYNOPSIS
    Generates standard reports from ConnectSecure data APIs.
    Uses: application_vulnerabilities, application_vulnerabilities_suppressed, external_asset_vulnerabilities.
    .PARAMETER VulnerabilityData
    Optional pre-fetched application_vulnerabilities - when provided, reused for All Vulns, Executive Summary, Pending EPSS (avoids duplicate API calls).
    .PARAMETER OnProgress
    Optional scriptblock: param($message) - called to update progress UI during conversion and Excel export.
    #>
    param(
        [string]$InternalReportType,
        [string]$OutputPath,
        [int]$CompanyId,
        [string]$ClientName,
        [string]$ScanDate,
        [array]$VulnerabilityData = $null,
        [int]$TopCount = 10,
        [scriptblock]$OnProgress = $null,
        [int]$DebugLimit = 0
    )
    $vulnDataToPass = if ($InternalReportType -in @('all-vulnerabilities','executive-summary','pending-epss')) { $VulnerabilityData } else { $null }
    $rt = $InternalReportType
    $t0 = 'pending-epss'
    $t1 = 'executive-summary'
    $t2 = 'all-vulnerabilities'
    $t3 = 'external-vulnerabilities'
    $t4 = 'suppressed-vulnerabilities'
    if ($rt -eq $t0) { New-PendingEPSSReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -VulnerabilityData $vulnDataToPass -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t1) { New-ExecutiveSummaryReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate -VulnerabilityData $vulnDataToPass -TopCount $TopCount -DebugLimit $DebugLimit }
    elseif ($rt -eq $t2) { New-AllVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -VulnerabilityData $vulnDataToPass -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t3) { New-ExternalVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t4) { New-SuppressedVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    else { throw ('Unknown report type: ' + $InternalReportType) }
}
