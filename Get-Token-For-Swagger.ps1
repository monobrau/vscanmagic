# Get ConnectSecure API Token for Swagger UI Authorization
# This script authenticates and outputs the token in the format needed for Swagger UI

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

Write-Host "=" -ForegroundColor Cyan
Write-Host "ConnectSecure API Token Generator for Swagger UI" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""

# Import ConnectSecure API functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$connectSecureApiPath = Join-Path $scriptPath "ConnectSecure-API.ps1"
if (Test-Path $connectSecureApiPath) {
    . $connectSecureApiPath
} else {
    Write-Host "Error: ConnectSecure-API.ps1 not found in $scriptPath" -ForegroundColor Red
    exit 1
}

# Clean inputs
$TenantName = $TenantName.Trim() -replace "`r`n|`r|`n", ""
$ClientId = $ClientId.Trim() -replace "`r`n|`r|`n", ""
$ClientSecret = $ClientSecret.Trim() -replace "`r`n|`r|`n", ""

Write-Host "Authenticating..." -ForegroundColor Yellow
$connected = Connect-ConnectSecureAPI -BaseUrl $BaseUrl `
                                      -TenantName $TenantName `
                                      -ClientId $ClientId `
                                      -ClientSecret $ClientSecret

if ($connected) {
    $token = $script:ConnectSecureConfig.AccessToken
    
    if ($token) {
        Write-Host ""
        Write-Host "=" -ForegroundColor Green
        Write-Host "✓ Authentication Successful!" -ForegroundColor Green
        Write-Host "=" -ForegroundColor Green
        Write-Host ""
        Write-Host "Token for Swagger UI Authorization:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Bearer $token" -ForegroundColor White -BackgroundColor DarkGray
        Write-Host ""
        Write-Host "Instructions:" -ForegroundColor Yellow
        Write-Host "1. Copy the token above (including 'Bearer ' prefix)" -ForegroundColor White
        Write-Host "2. Go to Swagger UI at: $BaseUrl/apidocs/" -ForegroundColor White
        Write-Host "3. Click the 'Authorize' button (lock icon)" -ForegroundColor White
        Write-Host "4. In the 'Value' field, paste: Bearer $($token.Substring(0, [Math]::Min(50, $token.Length)))..." -ForegroundColor White
        Write-Host "5. Click 'Authorize' button" -ForegroundColor White
        Write-Host "6. Click 'Close'" -ForegroundColor White
        Write-Host ""
        Write-Host "Note: The token will expire after some time. Re-run this script to get a new token." -ForegroundColor Gray
        Write-Host ""
        
        # Copy to clipboard if available
        try {
            $fullToken = "Bearer $token"
            Set-Clipboard -Value $fullToken
            Write-Host "✓ Token copied to clipboard!" -ForegroundColor Green
        } catch {
            Write-Host "Could not copy to clipboard automatically" -ForegroundColor Yellow
        }
    } else {
        Write-Host ""
        Write-Host "✗ Authentication succeeded but no access token was returned" -ForegroundColor Red
        Write-Host "This might indicate an API issue. Check the console output above for details." -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "✗ Authentication Failed" -ForegroundColor Red
    Write-Host "Check the error messages above for troubleshooting steps." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "1. Tenant name format (try your email address)" -ForegroundColor White
    Write-Host "2. API key permissions" -ForegroundColor White
    Write-Host "3. 'Failed to create customer' error - contact ConnectSecure support" -ForegroundColor White
}
