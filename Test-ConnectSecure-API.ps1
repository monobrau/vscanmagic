<#
.SYNOPSIS
    Comprehensive ConnectSecure API Testing Script
    
.DESCRIPTION
    This script tests ConnectSecure API authentication with various configurations
    to help identify the correct tenant name format and troubleshoot authentication issues.
    
.PARAMETER BaseUrl
    ConnectSecure API base URL (e.g., https://pod104.myconnectsecure.com)
    
.PARAMETER TenantName
    Tenant name to test (e.g., river-run or cknospe@river-run.com)
    
.PARAMETER ClientId
    ConnectSecure API Client ID
    
.PARAMETER ClientSecret
    ConnectSecure API Client Secret
    
.PARAMETER TestAllFormats
    Test multiple tenant name format variations
    
.PARAMETER Endpoint
    Specific endpoint to test: "w/authorize" or "auth/v4/token". If not specified, tests both.
    
.EXAMPLE
    .\Test-ConnectSecure-API.ps1 -BaseUrl "https://pod104.myconnectsecure.com" -TenantName "river-run" -ClientId "c113d502-e391-4bb1-8d70-53afc01215b8" -ClientSecret "your-secret"
    
.EXAMPLE
    .\Test-ConnectSecure-API.ps1 -BaseUrl "https://pod104.myconnectsecure.com" -TenantName "river-run" -ClientId "c113d502-e391-4bb1-8d70-53afc01215b8" -ClientSecret "your-secret" -TestAllFormats
    
.EXAMPLE
    .\Test-ConnectSecure-API.ps1 -BaseUrl "https://pod104.myconnectsecure.com" -TenantName "river-run" -ClientId "c113d502-e391-4bb1-8d70-53afc01215b8" -ClientSecret "your-secret" -Endpoint "auth/v4/token"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$BaseUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantName,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,
    
    [switch]$TestAllFormats,
    
    [ValidateSet("w/authorize", "auth/v4/token", "")]
    [string]$Endpoint = ""
)

# Import ConnectSecure API functions if available
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$connectSecureApiPath = Join-Path $scriptPath "ConnectSecure-API.ps1"

if (Test-Path $connectSecureApiPath) {
    Write-Host "Loading ConnectSecure-API.ps1..." -ForegroundColor Gray
    . $connectSecureApiPath
    $useModule = $true
} else {
    Write-Host "Warning: ConnectSecure-API.ps1 not found. Using inline authentication logic." -ForegroundColor Yellow
    $useModule = $false
}

function Write-TestLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Header')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Header' { 'Cyan' }
        default { 'White' }
    }
    
    $prefix = switch ($Level) {
        'Success' { '✓' }
        'Warning' { '⚠' }
        'Error' { '✗' }
        'Header' { '=' }
        default { '•' }
    }
    
    Write-Host "[$timestamp] $prefix $Message" -ForegroundColor $color
}

function Test-ConnectSecureAuth {
    param(
        [string]$BaseUrl,
        [string]$TenantName,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Endpoint = "w/authorize"
    )
    
    Write-TestLog "Testing Authentication (Endpoint: $Endpoint)" -Level Header
    Write-TestLog "Base URL: $BaseUrl"
    Write-TestLog "Tenant Name: $TenantName"
    Write-TestLog "Client ID: $ClientId"
    Write-TestLog "Client Secret: *** (Length: $($ClientSecret.Length) chars)"
    Write-Host ""
    
    # Clean inputs
    $TenantName = $TenantName.Trim() -replace "`r`n|`r|`n", ""
    $ClientId = $ClientId.Trim() -replace "`r`n|`r|`n", ""
    $ClientSecret = $ClientSecret.Trim() -replace "`r`n|`r|`n", ""
    
    # Prepare request based on endpoint type
    $authUrl = "$($BaseUrl.TrimEnd('/'))/$Endpoint"
    Write-TestLog "Request URL: $authUrl"
    
    $headers = @{
        'accept' = 'application/json'
        'Content-Type' = 'application/json'
    }
    
    $body = ""
    
    if ($Endpoint -eq "w/authorize") {
        # Method 1: Header-based auth (PowerShell method)
        $authString = "${TenantName}+${ClientId}:${ClientSecret}"
        Write-TestLog "Auth String Format: ${TenantName}+${ClientId}:<secret>"
        Write-TestLog "Auth String Length: $($authString.Length) characters"
        
        # Base64 encode
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
            $base64Auth = [System.Convert]::ToBase64String($bytes)
            Write-TestLog "Base64 Token Generated: $($base64Auth.Length) characters" -Level Success
            Write-TestLog "Base64 (first 80 chars): $($base64Auth.Substring(0, [Math]::Min(80, $base64Auth.Length)))..."
        } catch {
            Write-TestLog "Failed to encode auth string: $($_.Exception.Message)" -Level Error
            return $false
        }
        
        $headers['Client-Auth-Token'] = $base64Auth
        $body = ""
        
    } elseif ($Endpoint -eq "auth/v4/token") {
        # Method 2: JSON body auth (Python method)
        Write-TestLog "Using JSON body authentication method" -Level Info
        $body = @{
            clientid = $ClientId
            clientsecret = $ClientSecret
            tenantname = $TenantName
        } | ConvertTo-Json
        Write-TestLog "Request Body: $($body -replace '"clientsecret":"[^"]+"', '"clientsecret":"***"')"
    }
    
    Write-Host ""
    Write-TestLog "Sending POST request..." -Level Info
    
    try {
        $authUri = [System.Uri]$authUrl
        
        if ($body) {
            $response = Invoke-RestMethod -Uri $authUri -Method Post -Headers $headers -Body $body -ErrorAction Stop
        } else {
            $response = Invoke-RestMethod -Uri $authUri -Method Post -Headers $headers -Body "" -ErrorAction Stop
        }
        
        Write-Host ""
        Write-TestLog "Response received!" -Level Success
        Write-TestLog "Response Keys: $($response.PSObject.Properties.Name -join ', ')"
        
        # Check response
        if ($response.status -eq $false) {
            $errorMsg = if ($response.message) { $response.message } else { "Unknown error" }
            Write-TestLog "Status: false" -Level Error
            Write-TestLog "Message: $errorMsg" -Level Error
            
            # Analyze error
            if ($errorMsg -eq "Failed to authorize") {
                Write-Host ""
                Write-TestLog "ANALYSIS: Credentials rejected" -Level Error
                Write-TestLog "  • Tenant name format might be wrong" -Level Warning
                Write-TestLog "  • Client ID or Secret might be incorrect" -Level Warning
                Write-TestLog "  • API key might be inactive" -Level Warning
            } elseif ($errorMsg -eq "Failed to create customer") {
                Write-Host ""
                Write-TestLog "ANALYSIS: Authentication passed but API error occurred" -Level Warning
                Write-TestLog "  • Credentials appear correct (past authorization check)" -Level Success
                Write-TestLog "  • This is likely a ConnectSecure API-side issue" -Level Warning
                Write-TestLog "  • Contact ConnectSecure support" -Level Warning
            }
            
            Write-Host ""
            Write-TestLog "Full Response:" -Level Info
            $response | ConvertTo-Json -Depth 3 | Write-Host
            
            return $false
        }
        
        $token = $response.access_token
        if (-not $token) {
            $token = $response.token
        }
        
        if ($token) {
            Write-Host ""
            Write-TestLog "✓ AUTHENTICATION SUCCESSFUL!" -Level Success
            Write-TestLog "Access Token: $($token.Substring(0, [Math]::Min(50, $token.Length)))..."
            if ($response.user_id) {
                Write-TestLog "User ID: $($response.user_id)"
            }
            Write-Host ""
            Write-TestLog "Token for Swagger UI:" -Level Info
            Write-Host "Bearer $token" -ForegroundColor Cyan
            return $true
        } else {
            Write-Host ""
            Write-TestLog "No access_token in response" -Level Error
            Write-TestLog "Full Response:" -Level Info
            $response | ConvertTo-Json -Depth 3 | Write-Host
            return $false
        }
        
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        
        Write-Host ""
        Write-TestLog "Request Failed" -Level Error
        Write-TestLog "Error: $($_.Exception.Message)" -Level Error
        
        if ($statusCode) {
            Write-TestLog "HTTP Status Code: $statusCode" -Level Error
            
            if ($statusCode -eq 404) {
                Write-TestLog "Endpoint not found - this endpoint may not exist" -Level Error
            } elseif ($statusCode -eq 502) {
                Write-TestLog "502 Bad Gateway - Server is down or unreachable" -Level Error
                Write-TestLog "  • ConnectSecure server may be temporarily unavailable" -Level Warning
                Write-TestLog "  • Try again in a few minutes" -Level Warning
                Write-TestLog "  • Check ConnectSecure status page if available" -Level Warning
            } elseif ($statusCode -eq 503) {
                Write-TestLog "503 Service Unavailable - Server is overloaded or maintenance" -Level Error
            } elseif ($statusCode -eq 401) {
                Write-TestLog "Unauthorized - credentials rejected" -Level Error
            } elseif ($statusCode -eq 400) {
                Write-TestLog "Bad Request - check request format" -Level Error
            }
        }
        
        # Try to read error response
        try {
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorDetails = $reader.ReadToEnd()
                $reader.Close()
                if ($errorDetails) {
                    # Check if it's HTML (like Cloudflare error page)
                    if ($errorDetails -match '<html' -or $errorDetails -match '<!DOCTYPE') {
                        Write-TestLog "Received HTML error page (likely Cloudflare)" -Level Error
                        if ($errorDetails -match '502') {
                            Write-TestLog "Cloudflare 502 Bad Gateway detected" -Level Error
                        }
                    } else {
                        Write-TestLog "Error Response Body:" -Level Error
                        Write-Host $errorDetails -ForegroundColor Red
                    }
                }
            }
        } catch { }
        
        return $false
    }
}

# Main execution
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host "ConnectSecure API Test Script" -ForegroundColor Cyan
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""

# Test base URL first
Write-TestLog "Testing Base URL accessibility..." -Level Header
$endpointsToTest = @("w/authorize", "auth/v4/token")
$baseUrlAccessible = $false

foreach ($endpoint in $endpointsToTest) {
    try {
        $testUrl = "$($BaseUrl.TrimEnd('/'))/$endpoint"
        Write-TestLog "Testing endpoint: $endpoint" -Level Info
        $testResponse = Invoke-WebRequest -Uri $testUrl -Method POST -Headers @{'accept'='application/json'} -Body "" -TimeoutSec 5 -ErrorAction Stop
        Write-TestLog "Endpoint '$endpoint' is accessible (HTTP $($testResponse.StatusCode))" -Level Success
        $baseUrlAccessible = $true
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        
        if ($statusCode -eq 404) {
            Write-TestLog "Endpoint '$endpoint' returns 404 - endpoint not found" -Level Warning
        } elseif ($statusCode -eq 502) {
            Write-TestLog "Endpoint '$endpoint' returns 502 Bad Gateway - server issue" -Level Error
        } elseif ($statusCode -eq 400 -or $statusCode -eq 401) {
            Write-TestLog "Endpoint '$endpoint' exists (HTTP $statusCode - endpoint is valid)" -Level Success
            $baseUrlAccessible = $true
        } else {
            Write-TestLog "Endpoint '$endpoint' returned HTTP $statusCode" -Level Warning
        }
    }
}

if (-not $baseUrlAccessible) {
    Write-TestLog "WARNING: Base URL may be incorrect or server is down" -Level Warning
    Write-TestLog "Expected format: https://pod{number}.myconnectsecure.com" -Level Info
}

Write-Host ""

# Test with provided tenant name - try specified endpoint or both
Write-Host "=" * 70 -ForegroundColor Cyan
Write-TestLog "Testing Authentication Methods..." -Level Header
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""

$result = $false
$successfulEndpoint = $null

if ($Endpoint) {
    # Test only specified endpoint
    Write-Host ""
    Write-TestLog "=== Testing Endpoint: /$Endpoint ===" -Level Header
    $result = Test-ConnectSecureAuth -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret -Endpoint $Endpoint
    if ($result) {
        $successfulEndpoint = $Endpoint
    }
} else {
    # Try both endpoints
    # Try Method 1: /w/authorize (PowerShell method)
    Write-Host ""
    Write-TestLog "=== METHOD 1: /w/authorize (Header-based) ===" -Level Header
    $result1 = Test-ConnectSecureAuth -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret -Endpoint "w/authorize"

    if ($result1) {
        $result = $true
        $successfulEndpoint = "w/authorize"
    }

    # Try Method 2: /auth/v4/token (Python method)
    if (-not $result) {
        Write-Host ""
        Write-TestLog "=== METHOD 2: /auth/v4/token (JSON body-based) ===" -Level Header
        $result2 = Test-ConnectSecureAuth -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret -Endpoint "auth/v4/token"
        
        if ($result2) {
            $result = $true
            $successfulEndpoint = "auth/v4/token"
        }
    }
}

# If TestAllFormats is specified, test variations
if ($TestAllFormats -and -not $result) {
    Write-Host ""
    Write-Host "=" * 70 -ForegroundColor Cyan
    Write-TestLog "Testing Additional Tenant Name Formats..." -Level Header
    Write-Host "=" * 70 -ForegroundColor Cyan
    Write-Host ""
    
    # Extract potential variations
    $baseTenant = $TenantName
    
    # Test variations
    $variations = @()
    
    # If it's an email, try just the domain part
    if ($baseTenant -match '@') {
        $parts = $baseTenant -split '@'
        $variations += $parts[0]  # username part
        $variations += $parts[1]  # domain part
        $variations += $parts[1].Replace('.com', '').Replace('.', '-')  # domain without .com
    }
    
    # If it's not an email, try email format
    if ($baseTenant -notmatch '@') {
        $variations += "$baseTenant@river-run.com"
        $variations += "cknospe@$baseTenant.com"
    }
    
    # Try case variations
    $variations += $baseTenant.ToLower()
    $variations += $baseTenant.ToUpper()
    
    # Remove duplicates
    $variations = $variations | Where-Object { $_ -ne $baseTenant } | Select-Object -Unique
    
    foreach ($variant in $variations) {
        Write-Host ""
        Write-TestLog "Testing variant: $variant" -Level Header
        $variantResult = Test-ConnectSecureAuth -BaseUrl $BaseUrl -TenantName $variant -ClientId $ClientId -ClientSecret $ClientSecret
        
        if ($variantResult) {
            Write-Host ""
            Write-TestLog "SUCCESS! Correct tenant name format: $variant" -Level Success
            break
        }
    }
}

Write-Host ""
Write-Host "=" * 70 -ForegroundColor Cyan
Write-TestLog "Test Complete" -Level Header
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""

# Summary
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Cyan
Write-TestLog "Summary" -Level Header
Write-Host "=" * 70 -ForegroundColor Cyan
Write-TestLog "Base URL: $BaseUrl"
Write-TestLog "Tenant Name Tested: $TenantName"
if ($successfulEndpoint) {
    Write-TestLog "Successful Endpoint: $successfulEndpoint" -Level Success
}
Write-TestLog "Result: $(if ($result) { 'SUCCESS' } else { 'FAILED' })" -Level $(if ($result) { 'Success' } else { 'Error' })
Write-Host ""

if (-not $result) {
    Write-TestLog "Next Steps:" -Level Header
    Write-TestLog "1. If you see 502 errors, the server may be down - try again later" -Level Info
    Write-TestLog "2. Verify tenant name on API Key page matches exactly" -Level Info
    Write-TestLog "3. Check API key is Active in ConnectSecure portal" -Level Info
    Write-TestLog "4. Try running with -TestAllFormats to test tenant name variations" -Level Info
    Write-TestLog "5. Contact ConnectSecure support if issue persists" -Level Info
    Write-Host ""
}
