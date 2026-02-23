#Requires -Version 5.1
<#
.SYNOPSIS
Test script for VScanMagic API Server

.DESCRIPTION
This script demonstrates how to use the VScanMagic API endpoints.
Make sure the API server is running before executing this script.

.PARAMETER ApiUrl
Base URL of the API server (default: http://localhost:8080)

.PARAMETER InputFile
Path to the input Excel file for testing

.PARAMETER ClientName
Client name for reports (default: "Test Client")

.EXAMPLE
.\Test-API.ps1 -InputFile "C:\Reports\VulnerabilityScan.xlsx"
#>

param(
    [string]$ApiUrl = "http://localhost:8080",
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    [string]$ClientName = "Test Client"
)

Write-Host "VScanMagic API Test Script" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan
Write-Host ""

# Check if input file exists
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# Test health check
Write-Host "Testing health check endpoint..." -ForegroundColor Yellow
try {
    $health = Invoke-RestMethod -Uri "$ApiUrl/health" -Method GET
    Write-Host "✓ API Server is healthy" -ForegroundColor Green
    Write-Host "  Status: $($health.status)" -ForegroundColor Gray
    Write-Host "  Version: $($health.version)" -ForegroundColor Gray
    Write-Host ""
} catch {
    Write-Error "Failed to connect to API server at $ApiUrl"
    Write-Host "Make sure the API server is running: .\VScanMagic-API.ps1" -ForegroundColor Yellow
    exit 1
}

# Prepare request body
$requestBody = @{
    inputPath = $InputFile
    clientName = $ClientName
    scanDate = (Get-Date).ToString("MM/dd/yyyy")
} | ConvertTo-Json

# Test endpoints
$endpoints = @(
    @{ Name = "Pending EPSS Report"; Type = "pending-epss"; Extension = "xlsx" },
    @{ Name = "Executive Summary"; Type = "executive-summary"; Extension = "docx" },
    @{ Name = "All Vulnerabilities Report"; Type = "all-vulnerabilities"; Extension = "xlsx" },
    @{ Name = "External Vulnerabilities Report"; Type = "external-vulnerabilities"; Extension = "xlsx" },
    @{ Name = "Suppressed Vulnerabilities Report"; Type = "suppressed-vulnerabilities"; Extension = "xlsx" }
)

$outputDir = Join-Path $PSScriptRoot "API_Test_Output"
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

foreach ($endpoint in $endpoints) {
    Write-Host "Testing $($endpoint.Name)..." -ForegroundColor Yellow
    try {
        $outputFile = Join-Path $outputDir "$($endpoint.Name).$($endpoint.Extension)"
        $uri = "$ApiUrl/api/reports/$($endpoint.Type)"
        
        Invoke-RestMethod -Uri $uri `
            -Method POST `
            -Body $requestBody `
            -ContentType "application/json" `
            -OutFile $outputFile
        
        if (Test-Path $outputFile) {
            $fileSize = (Get-Item $outputFile).Length
            Write-Host "✓ Generated: $outputFile ($fileSize bytes)" -ForegroundColor Green
        } else {
            Write-Host "✗ File not created" -ForegroundColor Red
        }
    } catch {
        Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.ErrorDetails.Message) {
            try {
                $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json
                Write-Host "  Details: $($errorDetails.error)" -ForegroundColor Red
            } catch {
                Write-Host "  Details: $($_.ErrorDetails.Message)" -ForegroundColor Red
            }
        }
    }
    Write-Host ""
}

Write-Host "Test complete. Output files saved to: $outputDir" -ForegroundColor Cyan
