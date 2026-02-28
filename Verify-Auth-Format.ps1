# Verify ConnectSecure Authentication Format
# This script helps verify the exact format being used

param(
    [Parameter(Mandatory=$true)]
    [string]$TenantName,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret
)

Write-Host "=" -ForegroundColor Cyan
Write-Host "ConnectSecure Auth Format Verification" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""

# Clean inputs
$TenantName = $TenantName.Trim() -replace "`r`n|`r|`n", ""
$ClientId = $ClientId.Trim() -replace "`r`n|`r|`n", ""
$ClientSecret = $ClientSecret.Trim() -replace "`r`n|`r|`n", ""

# Construct auth string
$authString = "${TenantName}+${ClientId}:${ClientSecret}"

Write-Host "Input Values:" -ForegroundColor Yellow
Write-Host "  Tenant Name: '$TenantName' (Length: $($TenantName.Length))" -ForegroundColor White
Write-Host "  Client ID: '$ClientId' (Length: $($ClientId.Length))" -ForegroundColor White
Write-Host "  Client Secret: '***' (Length: $($ClientSecret.Length))" -ForegroundColor White
Write-Host ""

Write-Host "Auth String Format:" -ForegroundColor Yellow
Write-Host "  $TenantName+$ClientId:<secret>" -ForegroundColor White
Write-Host ""

Write-Host "Full Auth String (first 100 chars):" -ForegroundColor Yellow
Write-Host "  $($authString.Substring(0, [Math]::Min(100, $authString.Length)))..." -ForegroundColor White
Write-Host "  Total Length: $($authString.Length) characters" -ForegroundColor White
Write-Host ""

# Base64 encode
$bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$base64Auth = [System.Convert]::ToBase64String($bytes)

Write-Host "Base64 Encoded Token:" -ForegroundColor Yellow
Write-Host "  $base64Auth" -ForegroundColor White
Write-Host ""

# Decode to verify
Write-Host "Verification (decoding base64 back):" -ForegroundColor Yellow
try {
    $decodedBytes = [System.Convert]::FromBase64String($base64Auth)
    $decodedString = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
    
    Write-Host "  Decoded matches original: $($decodedString -eq $authString)" -ForegroundColor $(if ($decodedString -eq $authString) { "Green" } else { "Red" })
    Write-Host "  Decoded (first 100 chars): $($decodedString.Substring(0, [Math]::Min(100, $decodedString.Length)))..." -ForegroundColor White
    Write-Host ""
    
    # Parse the decoded string
    if ($decodedString -match "^(.+)\+(.+):(.+)$") {
        $parsedTenant = $matches[1]
        $parsedClientId = $matches[2]
        $parsedSecret = $matches[3]
        
        Write-Host "Parsed Components:" -ForegroundColor Yellow
        Write-Host "  Tenant: '$parsedTenant'" -ForegroundColor White
        Write-Host "  Client ID: '$parsedClientId'" -ForegroundColor White
        Write-Host "  Secret Length: $($parsedSecret.Length) chars" -ForegroundColor White
        Write-Host ""
        
        Write-Host "Verification:" -ForegroundColor Yellow
        Write-Host "  Tenant matches: $($parsedTenant -eq $TenantName)" -ForegroundColor $(if ($parsedTenant -eq $TenantName) { "Green" } else { "Red" })
        Write-Host "  Client ID matches: $($parsedClientId -eq $ClientId)" -ForegroundColor $(if ($parsedClientId -eq $ClientId) { "Green" } else { "Red" })
        Write-Host "  Secret length matches: $($parsedSecret.Length -eq $ClientSecret.Length)" -ForegroundColor $(if ($parsedSecret.Length -eq $ClientSecret.Length) { "Green" } else { "Red" })
    }
} catch {
    Write-Host "  ERROR: Could not decode base64: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=" -ForegroundColor Cyan
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Check the API Key page in ConnectSecure:" -ForegroundColor Yellow
Write-Host "   • Go to: Global > Settings > Users > [Your User] > Action > API Key" -ForegroundColor White
Write-Host "   • Look for a field labeled 'Tenant Name', 'Tenant', or 'Tenant ID'" -ForegroundColor White
Write-Host "   • Copy it EXACTLY as shown" -ForegroundColor White
Write-Host ""
Write-Host "2. Common tenant name formats:" -ForegroundColor Yellow
Write-Host "   • Email: user@company.com" -ForegroundColor White
Write-Host "   • Company ID: company-123" -ForegroundColor White
Write-Host "   • UUID: 12345678-1234-1234-1234-123456789abc" -ForegroundColor White
Write-Host "   • Organization name: YourOrgName" -ForegroundColor White
Write-Host ""
Write-Host "3. Current tenant name being used: '$TenantName'" -ForegroundColor Yellow
Write-Host "   If this doesn't match what's on the API Key page, update it!" -ForegroundColor White
Write-Host ""
