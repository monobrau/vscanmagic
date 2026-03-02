<#
.SYNOPSIS
    Gets or creates an OneDrive share link for a local file or folder path.
.DESCRIPTION
    Uses Microsoft Graph API to create a sharing link for files in OneDrive sync folders.
    Requires: Install-Module Microsoft.Graph -Scope CurrentUser
    On first run, you'll be prompted to sign in (browser).
.PARAMETER Path
    Full local path to the file or folder (e.g. K:\OneDrive\...\2026 - Q1)
.PARAMETER LinkType
    view = read-only, edit = read-write. Default: view
.PARAMETER Scope
    organization = only your org can access; anonymous = anyone with link. Default: organization
.EXAMPLE
    .\Get-OneDriveShareLink.ps1 -Path "K:\OneDrive\OneDrive - River Run Computers, Inc\General\Accurate Metal\Network Documentation\Vulnerability Scans\2026 - Q1"
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [ValidateSet('view','edit')]
    [string]$LinkType = 'view',
    [ValidateSet('organization','anonymous')]
    [string]$Scope = 'organization',
    [switch]$Quiet = $false
)

$ErrorActionPreference = 'Stop'

# Check for Microsoft.Graph
$graphModule = Get-Module -ListAvailable Microsoft.Graph
if (-not $graphModule) {
    Write-Error @"
Microsoft.Graph module is required. Install it with:
  Install-Module Microsoft.Graph -Scope CurrentUser

Then run this script again.
"@
}

# Ensure we have a valid path
if (-not (Test-Path $Path)) {
    Write-Error "Path not found: $Path"
}
$Path = (Resolve-Path $Path).Path

# Extract OneDrive-relative path: everything after "OneDrive - TenantName\"
# e.g. K:\OneDrive\OneDrive - River Run Computers, Inc\General\... -> General\...
$onedriveMatch = [regex]::Match($Path, '(?i)OneDrive\s*-\s*[^\\]+\\(.+)')
if (-not $onedriveMatch.Success) {
    Write-Error @"
Path does not appear to be in an OneDrive sync folder.
Expected format: ...\OneDrive - [TenantName]\[folder path]
Example: K:\OneDrive\OneDrive - River Run Computers, Inc\General\Client\2026 - Q1
"@
}
$relativePath = $onedriveMatch.Groups[1].Value -replace '\\', '/'
# Graph API: encode each path segment, keep / as separator
$pathForUri = ($relativePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'

# Connect to Graph (interactive sign-in if needed)
$requiredScopes = @('Files.ReadWrite', 'User.Read')
$context = Get-MgContext -ErrorAction SilentlyContinue
$needsConnect = (-not $context) -or ($context.Scopes -notcontains 'Files.ReadWrite')
if ($needsConnect) {
    Write-Host "Signing in to Microsoft Graph (browser will open)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
}

try {
    # Get item ID by path (Graph API: /drive/root:/path/to/item)
    $uri = "v1.0/me/drive/root:/$pathForUri"
    $item = Invoke-MgGraphRequest -Uri $uri -Method GET
    $itemId = $item.id

    # Create or get existing share link
    $body = @{
        type   = $LinkType
        scope  = $Scope
    } | ConvertTo-Json
    $linkUri = "v1.0/me/drive/items/$itemId/createLink"
    $result = Invoke-MgGraphRequest -Uri $linkUri -Method POST -Body $body -ContentType 'application/json'

    $webUrl = $result.link.webUrl
    if (-not $Quiet) { Write-Host "Share link: $webUrl" -ForegroundColor Green }
    return $webUrl
} catch {
    Write-Error "Failed to create share link: $($_.Exception.Message)"
}
