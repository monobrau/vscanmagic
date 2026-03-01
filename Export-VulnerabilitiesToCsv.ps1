<#
.SYNOPSIS
    Exports all vulnerabilities for a company to CSV for baseline application database.
.DESCRIPTION
    Pulls vulnerability data from an All Vulnerabilities Excel file or ConnectSecure API,
    then exports to CSV. Use -UniqueProductsOnly for a deduplicated application list
    suitable as a baseline for a remediation database (add Remediation Steps column for technicians).
.PARAMETER ExcelPath
    Path to All Vulnerabilities XLSX (from ConnectSecure download or VScanMagic). Use this OR ConnectSecure params.
.PARAMETER OutputPath
    Path for the output CSV. Default: .\VulnerabilitiesExport_<timestamp>.csv
.PARAMETER UniqueProductsOnly
    Output one row per unique Product/Application (aggregated). Ideal for building a remediation steps database.
.PARAMETER CompanyId
    ConnectSecure company ID. When set with API credentials, downloads All Vulnerabilities for that company first.
.PARAMETER ClientName
    Client/company name (for temp file naming when using ConnectSecure).
.PARAMETER ConnectSecureBaseUrl
    ConnectSecure API base URL (e.g. https://yourtenant.connectsecure.com).
.PARAMETER ConnectSecureTenant
    ConnectSecure tenant name.
.PARAMETER ConnectSecureClientId
    ConnectSecure API client ID.
.PARAMETER ConnectSecureClientSecret
    ConnectSecure API client secret.
.PARAMETER UseSavedCredentials
    Load ConnectSecure credentials from VScanMagic saved settings (requires CompanyId).
.EXAMPLE
    .\Export-VulnerabilitiesToCsv.ps1 -ExcelPath "C:\Reports\Client All Vulnerabilities.xlsx" -OutputPath ".\baseline.csv"
.EXAMPLE
    .\Export-VulnerabilitiesToCsv.ps1 -ExcelPath ".\All Vulnerabilities.xlsx" -UniqueProductsOnly -OutputPath ".\applications-baseline.csv"
.EXAMPLE
    .\Export-VulnerabilitiesToCsv.ps1 -CompanyId 123 -ClientName "Accurate Metal" -UseSavedCredentials -UniqueProductsOnly
#>
param(
    [string]$ExcelPath = "",
    [string]$OutputPath = "",
    [switch]$UniqueProductsOnly = $false,
    [int]$CompanyId = 0,
    [string]$ClientName = "Client",
    [string]$ConnectSecureBaseUrl = "",
    [string]$ConnectSecureTenant = "",
    [string]$ConnectSecureClientId = "",
    [string]$ConnectSecureClientSecret = "",
    [switch]$UseSavedCredentials = $false
)

$ErrorActionPreference = "Stop"

# Set script directory for module loading
$script:ScriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# Load VScanMagic modules (Core, Data)
$modulesDir = Join-Path $script:ScriptDirectory "VScanMagic-Modules"
if (-not (Test-Path $modulesDir)) {
    Write-Error "VScanMagic-Modules folder not found at: $modulesDir"
    exit 1
}
. (Join-Path $modulesDir "VScanMagic-Core.ps1")
. (Join-Path $modulesDir "VScanMagic-Data.ps1")

# Determine data source
$tempExcelPath = $null
$pathToRead = $ExcelPath

if ($CompanyId -gt 0) {
    # ConnectSecure path: need to download All Vulnerabilities first
    if (-not (Test-Path (Join-Path $script:ScriptDirectory "ConnectSecure-API.ps1"))) {
        Write-Error "ConnectSecure-API.ps1 not found. Cannot pull from ConnectSecure."
        exit 1
    }
    . (Join-Path $script:ScriptDirectory "ConnectSecure-API.ps1")

    # Load credentials
    if ($UseSavedCredentials) {
        $credsPath = Join-Path $env:LOCALAPPDATA "VScanMagic\ConnectSecure-Credentials.json"
        if (-not (Test-Path $credsPath)) {
            Write-Error "Saved ConnectSecure credentials not found at: $credsPath. Configure API in VScanMagic GUI first, or provide credentials explicitly."
            exit 1
        }
        $creds = Get-Content $credsPath -Raw | ConvertFrom-Json
        $ConnectSecureBaseUrl = $creds.BaseUrl
        $ConnectSecureTenant = $creds.TenantName
        $ConnectSecureClientId = $creds.ClientId
        $ConnectSecureClientSecret = $creds.ClientSecret
    }
    if ([string]::IsNullOrWhiteSpace($ConnectSecureBaseUrl) -or [string]::IsNullOrWhiteSpace($ConnectSecureTenant) -or [string]::IsNullOrWhiteSpace($ConnectSecureClientId) -or [string]::IsNullOrWhiteSpace($ConnectSecureClientSecret)) {
        Write-Error "ConnectSecure credentials required when using -CompanyId. Provide -UseSavedCredentials or -ConnectSecureBaseUrl, -ConnectSecureTenant, -ConnectSecureClientId, -ConnectSecureClientSecret."
        exit 1
    }

    Write-Host "Connecting to ConnectSecure..." -ForegroundColor Cyan
    $connected = Connect-ConnectSecureAPI -BaseUrl $ConnectSecureBaseUrl -TenantName $ConnectSecureTenant -ClientId $ConnectSecureClientId -ClientSecret $ConnectSecureClientSecret
    if (-not $connected) {
        Write-Error "Failed to authenticate with ConnectSecure."
        exit 1
    }

    $tempExcelPath = Join-Path $env:TEMP ("VScanMagic_Export_$([Guid]::NewGuid().ToString('N'))_AllVulnerabilities.xlsx")
    Write-Host "Downloading All Vulnerabilities for Company $CompanyId ($ClientName)..." -ForegroundColor Cyan
    try {
        New-AllVulnerabilitiesReportFromConnectSecure -OutputPath $tempExcelPath -CompanyId $CompanyId -ClientName $ClientName
    } catch {
        Write-Error "Failed to download from ConnectSecure: $($_.Exception.Message)"
        exit 1
    }
    $pathToRead = $tempExcelPath
}

if ([string]::IsNullOrWhiteSpace($pathToRead)) {
    Write-Error "Provide -ExcelPath (path to All Vulnerabilities XLSX) or -CompanyId with ConnectSecure credentials."
    exit 1
}

if (-not (Test-Path $pathToRead)) {
    Write-Error "File not found: $pathToRead"
    exit 1
}

# Read vulnerability data
Write-Host "Reading vulnerability data from: $pathToRead" -ForegroundColor Cyan
try {
    $vulnData = Get-VulnerabilityData -ExcelPath $pathToRead
} catch {
    Write-Error "Failed to read vulnerability data: $($_.Exception.Message)"
    if ($tempExcelPath -and (Test-Path $tempExcelPath)) { Remove-Item $tempExcelPath -Force -ErrorAction SilentlyContinue }
    exit 1
}

# Clean up temp file
if ($tempExcelPath -and (Test-Path $tempExcelPath)) {
    Remove-Item $tempExcelPath -Force -ErrorAction SilentlyContinue
}

if (-not $vulnData -or $vulnData.Count -eq 0) {
    Write-Warning "No vulnerability records found."
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        # Write empty CSV with headers
        $headers = if ($UniqueProductsOnly) {
            "Product", "Critical", "High", "Medium", "Low", "Vulnerability Count", "Affected Hosts", "Remediation Steps"
        } else {
            "Host Name", "IP", "Username", "Product", "Critical", "High", "Medium", "Low", "Vulnerability Count", "EPSS Score"
        }
        $headers -join "," | Out-File -FilePath $OutputPath -Encoding UTF8
    }
    exit 0
}

# Build output data
if ($UniqueProductsOnly) {
    # Aggregate by Product for baseline application database
    $byProduct = $vulnData | Group-Object -Property Product
    $exportData = $byProduct | ForEach-Object {
        $g = $_.Group
        $critical = ($g | Measure-Object -Property Critical -Sum).Sum
        $high = ($g | Measure-Object -Property High -Sum).Sum
        $medium = ($g | Measure-Object -Property Medium -Sum).Sum
        $low = ($g | Measure-Object -Property Low -Sum).Sum
        $vulnCount = ($g | Measure-Object -Property 'Vulnerability Count' -Sum).Sum
        $hostCount = ($g | Group-Object -Property { "$($_.'Host Name')|$($_.IP)" }).Count
        [PSCustomObject]@{
            Product = $_.Name
            Critical = $critical
            High = $high
            Medium = $medium
            Low = $low
            'Vulnerability Count' = $vulnCount
            'Affected Hosts' = $hostCount
            'Remediation Steps' = ""  # Blank for technician to fill in
        }
    } | Sort-Object -Property 'Vulnerability Count' -Descending
    Write-Host "Aggregated to $($exportData.Count) unique applications." -ForegroundColor Green
} else {
    $exportData = $vulnData
    Write-Host "Exporting $($exportData.Count) vulnerability records." -ForegroundColor Green
}

# Output path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $OutputPath = Join-Path $script:ScriptDirectory "VulnerabilitiesExport_$timestamp.csv"
}

# Export to CSV
$exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported to: $OutputPath" -ForegroundColor Green
