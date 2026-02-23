<#
.SYNOPSIS
    Test ConnectSecure API Endpoints - PowerShell equivalent of connectsecure_test.py
    
.DESCRIPTION
    Tests various ConnectSecure API endpoints after authentication, including:
    - Companies
    - Agents
    - Assets
    - Company Stats
    - Users
    
.PARAMETER BaseUrl
    ConnectSecure API base URL (e.g., https://pod104.myconnectsecure.com)
    
.PARAMETER TenantName
    Tenant name (e.g., river-run)
    
.PARAMETER ClientId
    ConnectSecure API Client ID
    
.PARAMETER ClientSecret
    ConnectSecure API Client Secret
    
.EXAMPLE
    .\Test-ConnectSecure-Endpoints.ps1 -BaseUrl "https://pod104.myconnectsecure.com" -TenantName "river-run" -ClientId "your-id" -ClientSecret "your-secret"
#>

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

# Import ConnectSecure API module if available
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$connectSecureApiPath = Join-Path $scriptPath "ConnectSecure-API.ps1"

if (Test-Path $connectSecureApiPath) {
    Write-Host "Loading ConnectSecure-API.ps1..." -ForegroundColor Gray
    . $connectSecureApiPath
    $useModule = $true
} else {
    Write-Host "Warning: ConnectSecure-API.ps1 not found. Using inline authentication." -ForegroundColor Yellow
    $useModule = $false
}

function Write-TestLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Header')]
        [string]$Level = 'Info'
    )
    
    $prefix = switch ($Level) {
        'Success' { '[+]' }
        'Warning' { '[!]' }
        'Error' { '[-]' }
        'Header' { '[*]' }
        default { '[*]' }
    }
    
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Header' { 'Cyan' }
        default { 'White' }
    }
    
    Write-Host "$prefix $Message" -ForegroundColor $color
}

function Get-AuthToken {
    param(
        [string]$BaseUrl,
        [string]$TenantName,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    Write-TestLog "Authenticating..." -Level Header
    
    if ($useModule) {
        # Use the module function
        $connected = Connect-ConnectSecureAPI -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret
        if ($connected) {
            # Verify token is stored
            if ($script:ConnectSecureConfig.AccessToken) {
                Write-TestLog "Authentication successful! User ID: $($script:ConnectSecureConfig.UserId)" -Level Success
                Write-Host ""
                return $script:ConnectSecureConfig.AccessToken
            } else {
                Write-TestLog "Authentication succeeded but token not stored" -Level Error
                return $null
            }
        }
        return $null
    } else {
        # Inline authentication
        $TenantName = $TenantName.Trim() -replace "`r`n|`r|`n", ""
        $ClientId = $ClientId.Trim() -replace "`r`n|`r|`n", ""
        $ClientSecret = $ClientSecret.Trim() -replace "`r`n|`r|`n", ""
        
        $authString = "${TenantName}+${ClientId}:${ClientSecret}"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
        $base64Auth = [System.Convert]::ToBase64String($bytes)
        
        $headers = @{
            'Client-Auth-Token' = $base64Auth
            'Content-Type' = 'application/json'
        }
        
        try {
            $authUrl = "$($BaseUrl.TrimEnd('/'))/w/authorize"
            $authUri = [System.Uri]$authUrl
            $response = Invoke-RestMethod -Uri $authUri -Method Post -Headers $headers -Body "" -ErrorAction Stop
            
            if (-not $response.status) {
                Write-TestLog "Auth failed: $($response | ConvertTo-Json -Depth 3)" -Level Error
                return $null
            }
            
            $token = $response.data.access_token
            if (-not $token) {
                $token = $response.access_token
            }
            
            Write-TestLog "Authentication successful! User ID: $($response.data.user_id)" -Level Success
            Write-Host ""
            return $token
        } catch {
            Write-TestLog "Authentication error: $($_.Exception.Message)" -Level Error
            return $null
        }
    }
}

function Test-GetCompanies {
    param(
        [string]$BaseUrl,
        [string]$Token
    )
    
    Write-TestLog "Fetching companies..." -Level Header
    
    try {
        if ($useModule) {
            # Use module's function
            $response = Invoke-ConnectSecureRequest -Endpoint "/r/company/companies" -Method "GET"
        } else {
            # Manual request
            $headers = @{
                'Authorization' = "Bearer $Token"
                'Content-Type' = 'application/json'
            }
            
            $url = "$($BaseUrl.TrimEnd('/'))/r/company/companies"
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        }
        
        $companies = if ($response -is [array]) { $response } else { $response.data }
        if (-not $companies) { $companies = @() }
        
        Write-TestLog "Companies found: $($companies.Count)" -Level Success
        if ($companies.Count -gt 0) {
            foreach ($c in $companies[0..([Math]::Min(4, $companies.Count - 1))]) {
                Write-Host "    - $($c.name) (ID: $($c.id))" -ForegroundColor Gray
            }
        }
        Write-Host ""
        return $companies
    } catch {
        Write-TestLog "Error fetching companies: $($_.Exception.Message)" -Level Error
        Write-Host ""
        return @()
    }
}

function Test-GetAgents {
    param(
        [string]$BaseUrl,
        [string]$Token
    )
    
    Write-TestLog "Fetching agents..." -Level Header
    
    try {
        if ($useModule) {
            $response = Invoke-ConnectSecureRequest -Endpoint "/r/company/agents" -Method "GET"
        } else {
            $headers = @{
                'Authorization' = "Bearer $Token"
                'Content-Type' = 'application/json'
            }
            
            $url = "$($BaseUrl.TrimEnd('/'))/r/company/agents"
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        }
        
        $agents = if ($response -is [array]) { $response } else { $response.data }
        if (-not $agents) { $agents = @() }
        
        Write-TestLog "Agents found: $($agents.Count)" -Level Success
        if ($agents.Count -gt 0) {
            Write-Host "    First agent: $($agents[0].hostname)" -ForegroundColor Gray
        }
        Write-Host ""
    } catch {
        Write-TestLog "Error fetching agents: $($_.Exception.Message)" -Level Error
        Write-Host ""
    }
}

function Test-GetAssets {
    param(
        [string]$BaseUrl,
        [string]$Token
    )
    
    Write-TestLog "Fetching assets..." -Level Header
    
    try {
        if ($useModule) {
            $response = Invoke-ConnectSecureRequest -Endpoint "/r/asset/assets" -Method "GET"
        } else {
            $headers = @{
                'Authorization' = "Bearer $Token"
                'Content-Type' = 'application/json'
            }
            
            $url = "$($BaseUrl.TrimEnd('/'))/r/asset/assets"
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        }
        
        $assets = if ($response -is [array]) { $response } else { $response.data }
        if (-not $assets) { $assets = @() }
        
        Write-TestLog "Assets found: $($assets.Count)" -Level Success
        if ($assets.Count -gt 0) {
            Write-Host "    First asset: $($assets[0].hostname)" -ForegroundColor Gray
        }
        Write-Host ""
    } catch {
        Write-TestLog "Error fetching assets: $($_.Exception.Message)" -Level Error
        Write-Host ""
    }
}

function Test-GetCompanyStats {
    param(
        [string]$BaseUrl,
        [string]$Token
    )
    
    Write-TestLog "Fetching company stats..." -Level Header
    
    try {
        if ($useModule) {
            $response = Invoke-ConnectSecureRequest -Endpoint "/r/company/company_stats" -Method "GET"
        } else {
            $headers = @{
                'Authorization' = "Bearer $Token"
                'Content-Type' = 'application/json'
            }
            
            $url = "$($BaseUrl.TrimEnd('/'))/r/company/company_stats"
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        }
        
        $stats = if ($response -is [array]) { $response } else { $response.data }
        if (-not $stats) { $stats = @() }
        
        Write-TestLog "Company stats records: $($stats.Count)" -Level Success
        Write-Host ""
    } catch {
        Write-TestLog "Error fetching company stats: $($_.Exception.Message)" -Level Error
        Write-Host ""
    }
}

function Test-GetUsers {
    param(
        [string]$BaseUrl,
        [string]$Token
    )
    
    Write-TestLog "Fetching users..." -Level Header
    
    try {
        if ($useModule) {
            $response = Invoke-ConnectSecureRequest -Endpoint "/r/user/get_users" -Method "GET"
        } else {
            $headers = @{
                'Authorization' = "Bearer $Token"
                'Content-Type' = 'application/json'
            }
            
            $url = "$($BaseUrl.TrimEnd('/'))/r/user/get_users"
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        }
        
        $users = if ($response -is [array]) { $response } else { $response.data }
        if (-not $users) { $users = @() }
        
        Write-TestLog "Users found: $($users.Count)" -Level Success
        Write-Host ""
    } catch {
        Write-TestLog "Error fetching users: $($_.Exception.Message)" -Level Error
        Write-Host ""
    }
}

# Main execution
Write-Host ""
Write-Host "=" * 50 -ForegroundColor Cyan
Write-Host "  ConnectSecure V4 API Test Script" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan
Write-Host ""

try {
    $token = Get-AuthToken -BaseUrl $BaseUrl -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret
    
    if (-not $token) {
        Write-TestLog "Failed to authenticate. Exiting." -Level Error
        exit 1
    }
    
    $companies = Test-GetCompanies -BaseUrl $BaseUrl -Token $token
    Test-GetAgents -BaseUrl $BaseUrl -Token $token
    Test-GetAssets -BaseUrl $BaseUrl -Token $token
    Test-GetCompanyStats -BaseUrl $BaseUrl -Token $token
    Test-GetUsers -BaseUrl $BaseUrl -Token $token
    
    Write-TestLog "All tests completed successfully." -Level Success
    
} catch {
    Write-TestLog "Unexpected error: $($_.Exception.Message)" -Level Error
    exit 1
}
