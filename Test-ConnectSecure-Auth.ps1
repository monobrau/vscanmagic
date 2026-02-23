# Test ConnectSecure API Authentication Manually
# This script helps debug authentication issues by testing the exact request

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
Write-Host "ConnectSecure API Authentication Test" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""

# Clean inputs
$TenantName = $TenantName.Trim() -replace "`r`n|`r|`n", ""
$ClientId = $ClientId.Trim() -replace "`r`n|`r|`n", ""
$ClientSecret = $ClientSecret.Trim() -replace "`r`n|`r|`n", ""

# Construct auth string
$authString = "${TenantName}+${ClientId}:${ClientSecret}"
Write-Host "Auth String Format: ${TenantName}+${ClientId}:<secret>" -ForegroundColor Yellow
Write-Host "Auth String Length: $($authString.Length) characters" -ForegroundColor Yellow
Write-Host ""

# Base64 encode
$bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$base64Auth = [System.Convert]::ToBase64String($bytes)
Write-Host "Base64 Token (first 100 chars): $($base64Auth.Substring(0, [Math]::Min(100, $base64Auth.Length)))..." -ForegroundColor Yellow
Write-Host "Base64 Token Length: $($base64Auth.Length) characters" -ForegroundColor Yellow
Write-Host ""

# Prepare request
$authUrl = "$($BaseUrl.TrimEnd('/'))/w/authorize"
Write-Host "Request URL: $authUrl" -ForegroundColor Green
Write-Host ""

$headers = @{
    'accept' = 'application/json'
    'Content-Type' = 'application/json'
    'Client-Auth-Token' = $base64Auth
}

Write-Host "Headers:" -ForegroundColor Green
$headers.GetEnumerator() | ForEach-Object {
    if ($_.Key -eq 'Client-Auth-Token') {
        Write-Host "  $($_.Key): $($_.Value.Substring(0, [Math]::Min(80, $_.Value.Length)))..." -ForegroundColor Gray
    } else {
        Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
    }
}
Write-Host ""

# Test 1: Using Invoke-RestMethod (what our code uses)
Write-Host "Test 1: Using Invoke-RestMethod..." -ForegroundColor Cyan
try {
    $authUri = [System.Uri]$authUrl
    $response = Invoke-RestMethod -Uri $authUri -Method Post -Headers $headers -Body "" -ErrorAction Stop
    
    Write-Host "✓ SUCCESS!" -ForegroundColor Green
    Write-Host "Response:" -ForegroundColor Green
    $response | ConvertTo-Json -Depth 3 | Write-Host
    
    if ($response.access_token) {
        Write-Host ""
        Write-Host "Access Token: $($response.access_token.Substring(0, [Math]::Min(50, $response.access_token.Length)))..." -ForegroundColor Green
        Write-Host "User ID: $($response.user_id)" -ForegroundColor Green
    }
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "✗ FAILED - HTTP $statusCode" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    
    try {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errorDetails = $reader.ReadToEnd()
        $reader.Close()
        Write-Host "Response Body:" -ForegroundColor Red
        Write-Host $errorDetails -ForegroundColor Red
    } catch {
        Write-Host "Could not read error response body" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=" -ForegroundColor Cyan
Write-Host "Test 2: Using Invoke-WebRequest (alternative method)..." -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""

# Test 2: Using Invoke-WebRequest
try {
    $response = Invoke-WebRequest -Uri $authUrl -Method Post -Headers $headers -Body "" -ErrorAction Stop
    
    Write-Host "✓ SUCCESS!" -ForegroundColor Green
    Write-Host "Status Code: $($response.StatusCode)" -ForegroundColor Green
    Write-Host "Response:" -ForegroundColor Green
    $response.Content | ConvertFrom-Json | ConvertTo-Json -Depth 3 | Write-Host
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "✗ FAILED - HTTP $statusCode" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    
    try {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errorDetails = $reader.ReadToEnd()
        $reader.Close()
        Write-Host "Response Body:" -ForegroundColor Red
        Write-Host $errorDetails -ForegroundColor Red
    } catch {
        Write-Host "Could not read error response body" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=" -ForegroundColor Cyan
Write-Host "Manual cURL Command:" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""
Write-Host "curl -X POST `"$authUrl`" \" -ForegroundColor Yellow
Write-Host "  -H `"accept: application/json`" \" -ForegroundColor Yellow
Write-Host "  -H `"Content-Type: application/json`" \" -ForegroundColor Yellow
Write-Host "  -H `"Client-Auth-Token: $base64Auth`"" -ForegroundColor Yellow
Write-Host ""
Write-Host "=" -ForegroundColor Cyan
Write-Host "Troubleshooting:" -ForegroundColor Cyan
Write-Host "=" -ForegroundColor Cyan
Write-Host ""
Write-Host "If both tests fail with 'Failed to authorize':" -ForegroundColor Yellow
Write-Host "1. Verify tenant name on API Key page matches: '$TenantName'" -ForegroundColor Yellow
Write-Host "2. Verify Client ID matches: '$ClientId'" -ForegroundColor Yellow
Write-Host "3. Verify Client Secret is correct (length: $($ClientSecret.Length) chars)" -ForegroundColor Yellow
Write-Host "4. Check if API key is ACTIVE in ConnectSecure portal" -ForegroundColor Yellow
Write-Host "5. Try the cURL command above in a terminal" -ForegroundColor Yellow
Write-Host ""
