#Requires -Modules Microsoft.PowerShell.Utility
<#
.SYNOPSIS
VScanMagic API Server - REST API for generating and downloading vulnerability reports

.DESCRIPTION
This script provides a REST API server that allows triggering and downloading:
- Pending EPSS Report (XLSX)
- Executive Summary (DOCX)
- All Vulnerabilities Report (XLSX)
- External Vulnerabilities Report (XLSX)
- Suppressed Vulnerabilities Report (XLSX)

.NOTES
Version: 1.0.0
Requires: Microsoft Excel and Microsoft Word installed.
Author: River Run MSP
#>

# --- Add Required Assemblies ---
# HttpListener is in System.dll (not a separate assembly)
Add-Type -AssemblyName System
Add-Type -AssemblyName System.Web

# --- Configuration ---
$script:ApiConfig = @{
    Port = 8080
    Host = "localhost"
    TempDirectory = Join-Path $env:TEMP "VScanMagic-API"
}

# --- Helper: Define Write-ApiLog before dot-sourcing (it is used below) ---
function Write-ApiLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}

# --- Import Functions from VScanMagic-GUI.ps1 ---
# Note: This assumes VScanMagic-GUI.ps1 is in the same directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$guiScriptPath = Join-Path $scriptPath "VScanMagic-GUI.ps1"
$connectSecurePath = Join-Path $scriptPath "ConnectSecure-API.ps1"

if (Test-Path $guiScriptPath) {
    Write-ApiLog "Loading functions from VScanMagic-GUI.ps1..." -Level Info
    
    # Set a flag BEFORE dot-sourcing to prevent GUI from starting
    # This must be in script scope to match what VScanMagic-GUI.ps1 checks
    $script:IsApiMode = $true
    
    # Dot-source the script to import functions
    # The GUI script will check $script:IsApiMode and skip GUI initialization if true
    . $guiScriptPath
} else {
    Write-ApiLog "VScanMagic-GUI.ps1 not found at: $guiScriptPath" -Level Error
    exit 1
}

# --- Import ConnectSecure API Functions ---
if (Test-Path $connectSecurePath) {
    Write-ApiLog "Loading ConnectSecure API functions..." -Level Info
    . $connectSecurePath
} else {
    Write-ApiLog "ConnectSecure-API.ps1 not found. ConnectSecure integration will be unavailable." -Level Warning
}

# --- Helper Functions ---

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerContext]$Context,
        [object]$Data,
        [int]$StatusCode = 200
    )

    $response = $Context.Response
    $response.StatusCode = $StatusCode
    $response.ContentType = "application/json"
    
    $json = $Data | ConvertTo-Json -Depth 10
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

function Send-FileResponse {
    param(
        [System.Net.HttpListenerContext]$Context,
        [string]$FilePath,
        [string]$ContentType = "application/octet-stream"
    )

    if (-not (Test-Path $FilePath)) {
        Send-JsonResponse -Context $Context -Data @{ error = "File not found: $FilePath" } -StatusCode 404
        return
    }

    $response = $Context.Response
    $response.StatusCode = 200
    $response.ContentType = $ContentType
    
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $response.Headers.Add("Content-Disposition", "attachment; filename=`"$fileName`"")
    
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $response.ContentLength64 = $fileBytes.Length
    $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
    $response.OutputStream.Close()
    
    Write-ApiLog "Sent file: $fileName ($($fileBytes.Length) bytes)" -Level Success
}

function Parse-RequestBody {
    param(
        [System.Net.HttpListenerRequest]$Request
    )

    $reader = New-Object System.IO.StreamReader($Request.InputStream)
    $body = $reader.ReadToEnd()
    $reader.Close()
    
    if ([string]::IsNullOrWhiteSpace($body)) {
        return @{}
    }
    
    try {
        return $body | ConvertFrom-Json
    } catch {
        Write-ApiLog "Failed to parse request body: $($_.Exception.Message)" -Level Error
        return @{}
    }
}

function Get-QueryParameter {
    param(
        [System.Net.HttpListenerRequest]$Request,
        [string]$ParameterName
    )

    $query = $Request.QueryString[$ParameterName]
    return $query
}

# --- Report Generation Functions ---

function New-AllVulnerabilitiesReport {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    Write-ApiLog "Generating All Vulnerabilities Report..."

    $excel = $null
    $workbook = $null
    $outputWorkbook = $null

    try {
        # Create Excel COM Object
        $excel = New-Object -ComObject Excel.Application
        if ($null -eq $excel) {
            throw "Failed to create Excel COM object. Make sure Microsoft Excel is installed."
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # Open input workbook
        if (Test-FileLocked $InputPath) {
            throw "The file is in use by another process. Please close it and try again."
        }
        $workbook = $excel.Workbooks.Open($InputPath)
        if ($null -eq $workbook) {
            throw "Failed to open workbook. File may be corrupted or in use."
        }

        # Find all vulnerability sheets (Critical, High, Medium, Low)
        $vulnerabilitySheets = @()
        $sheetPatterns = @("*Critical Vulnerabilities*", "*High Vulnerabilities*", "*Medium Vulnerabilities*", "*Low Vulnerabilities*")
        
        foreach ($sheet in $workbook.Worksheets) {
            $sheetName = $sheet.Name
            foreach ($pattern in $sheetPatterns) {
                if ($sheetName -like $pattern) {
                    $vulnerabilitySheets += $sheet
                    Write-ApiLog "Found vulnerability sheet: $sheetName"
                    break
                }
            }
        }

        if ($vulnerabilitySheets.Count -eq 0) {
            throw "No vulnerability sheets found in workbook"
        }

        # Create new workbook for output
        $outputWorkbook = $excel.Workbooks.Add()
        
        # Copy each vulnerability sheet to output workbook
        foreach ($sourceSheet in $vulnerabilitySheets) {
            $sourceSheet.Copy([System.Reflection.Missing]::Value, $outputWorkbook.Worksheets[$outputWorkbook.Worksheets.Count])
            Write-ApiLog "Copied sheet: $($sourceSheet.Name)"
        }

        # Delete default Sheet1 if it exists
        try {
            $defaultSheet = $outputWorkbook.Worksheets.Item(1)
            if ($defaultSheet.Name -eq "Sheet1") {
                $defaultSheet.Delete()
            }
            Clear-ComObject $defaultSheet
        } catch { }

        # Close source workbook
        $workbook.Close($false)
        Clear-ComObject $workbook
        $workbook = $null

        # Save output workbook
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $outputWorkbook.SaveAs($OutputPath)
        $outputWorkbook.Close($false)
        Clear-ComObject $outputWorkbook
        $outputWorkbook = $null

        Write-ApiLog "All Vulnerabilities Report saved to: $OutputPath" -Level Success

    } catch {
        Write-ApiLog "Error generating All Vulnerabilities Report: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($outputWorkbook) {
            try {
                $outputWorkbook.Close($false)
                Clear-ComObject $outputWorkbook
            } catch { }
        }
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

function New-ExternalVulnerabilitiesReport {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    Write-ApiLog "Generating External Vulnerabilities Report..."

    $excel = $null
    $workbook = $null
    $outputWorkbook = $null

    try {
        # Create Excel COM Object
        $excel = New-Object -ComObject Excel.Application
        if ($null -eq $excel) {
            throw "Failed to create Excel COM object. Make sure Microsoft Excel is installed."
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # Open input workbook
        if (Test-FileLocked $InputPath) {
            throw "The file is in use by another process. Please close it and try again."
        }
        $workbook = $excel.Workbooks.Open($InputPath)
        if ($null -eq $workbook) {
            throw "Failed to open workbook. File may be corrupted or in use."
        }

        # Find External Scan sheet
        $externalSheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -like "*External*" -or $sheet.Name -like "*External Scan*") {
                $externalSheet = $sheet
                Write-ApiLog "Found External sheet: $($sheet.Name)"
                break
            }
        }

        if ($null -eq $externalSheet) {
            throw "External Scan sheet not found in workbook"
        }

        # Create new workbook and copy external sheet
        $outputWorkbook = $excel.Workbooks.Add()
        $externalSheet.Copy([System.Reflection.Missing]::Value, $outputWorkbook.Worksheets[$outputWorkbook.Worksheets.Count])
        
        # Delete default Sheet1
        try {
            $defaultSheet = $outputWorkbook.Worksheets.Item(1)
            if ($defaultSheet.Name -eq "Sheet1") {
                $defaultSheet.Delete()
            }
            Clear-ComObject $defaultSheet
        } catch { }

        # Close source workbook
        $workbook.Close($false)
        Clear-ComObject $workbook
        $workbook = $null

        # Save output workbook
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $outputWorkbook.SaveAs($OutputPath)
        $outputWorkbook.Close($false)
        Clear-ComObject $outputWorkbook
        $outputWorkbook = $null

        Write-ApiLog "External Vulnerabilities Report saved to: $OutputPath" -Level Success

    } catch {
        Write-ApiLog "Error generating External Vulnerabilities Report: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($outputWorkbook) {
            try {
                $outputWorkbook.Close($false)
                Clear-ComObject $outputWorkbook
            } catch { }
        }
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

function New-SuppressedVulnerabilitiesReport {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    Write-ApiLog "Generating Suppressed Vulnerabilities Report..."

    $excel = $null
    $workbook = $null
    $outputWorkbook = $null

    try {
        # Create Excel COM Object
        $excel = New-Object -ComObject Excel.Application
        if ($null -eq $excel) {
            throw "Failed to create Excel COM object. Make sure Microsoft Excel is installed."
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # Open input workbook
        if (Test-FileLocked $InputPath) {
            throw "The file is in use by another process. Please close it and try again."
        }
        $workbook = $excel.Workbooks.Open($InputPath)
        if ($null -eq $workbook) {
            throw "Failed to open workbook. File may be corrupted or in use."
        }

        # Find Suppressed sheet
        $suppressedSheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -like "*Suppressed*" -or $sheet.Name -like "*Suppress*") {
                $suppressedSheet = $sheet
                Write-ApiLog "Found Suppressed sheet: $($sheet.Name)"
                break
            }
        }

        if ($null -eq $suppressedSheet) {
            throw "Suppressed Vulnerabilities sheet not found in workbook"
        }

        # Create new workbook and copy suppressed sheet
        $outputWorkbook = $excel.Workbooks.Add()
        $suppressedSheet.Copy([System.Reflection.Missing]::Value, $outputWorkbook.Worksheets[$outputWorkbook.Worksheets.Count])
        
        # Delete default Sheet1
        try {
            $defaultSheet = $outputWorkbook.Worksheets.Item(1)
            if ($defaultSheet.Name -eq "Sheet1") {
                $defaultSheet.Delete()
            }
            Clear-ComObject $defaultSheet
        } catch { }

        # Close source workbook
        $workbook.Close($false)
        Clear-ComObject $workbook
        $workbook = $null

        # Save output workbook
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $outputWorkbook.SaveAs($OutputPath)
        $outputWorkbook.Close($false)
        Clear-ComObject $outputWorkbook
        $outputWorkbook = $null

        Write-ApiLog "Suppressed Vulnerabilities Report saved to: $OutputPath" -Level Success

    } catch {
        Write-ApiLog "Error generating Suppressed Vulnerabilities Report: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($outputWorkbook) {
            try {
                $outputWorkbook.Close($false)
                Clear-ComObject $outputWorkbook
            } catch { }
        }
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

function New-ExecutiveSummaryReport {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$ClientName,
        [string]$ScanDate
    )

    Write-ApiLog "Generating Executive Summary Report..."

    try {
        # Read vulnerability data
        $vulnData = Get-VulnerabilityData -ExcelPath $InputPath
        
        if ($null -eq $vulnData -or $vulnData.Count -eq 0) {
            throw "No vulnerability data found"
        }

        # Calculate top vulnerabilities (use top 10 for executive summary)
        $top10 = Get-Top10Vulnerabilities -VulnData $vulnData

        # Generate Word report with executive summary
        New-WordReport -OutputPath $OutputPath `
                      -ClientName $ClientName `
                      -ScanDate $ScanDate `
                      -Top10Data $top10 `
                      -TimeEstimates $null `
                      -IsRMITPlus $false `
                      -GeneralRecommendations @() `
                      -ReportTitle "Executive Summary"

        Write-ApiLog "Executive Summary Report saved to: $OutputPath" -Level Success

    } catch {
        Write-ApiLog "Error generating Executive Summary Report: $($_.Exception.Message)" -Level Error
        throw
    }
}

# --- ConnectSecure report functions are in ConnectSecure-API.ps1 (loaded above) ---

# --- API Request Handlers ---

function Handle-ReportRequest {
    param(
        [System.Net.HttpListenerContext]$Context,
        [string]$ReportType
    )

    $request = $Context.Request

    try {
        # Parse request body
        $body = Parse-RequestBody -Request $request
        
        # Get parameters - support both file-based and ConnectSecure-based
        $inputPath = if ($body.InputPath) { $body.InputPath } else { Get-QueryParameter -Request $request -ParameterName "inputPath" }
        $clientName = if ($body.ClientName) { $body.ClientName } else { Get-QueryParameter -Request $request -ParameterName "clientName" }
        $scanDate = if ($body.ScanDate) { $body.ScanDate } else { Get-QueryParameter -Request $request -ParameterName "scanDate" }
        
        # ConnectSecure API parameters
        $useConnectSecure = if ($body.UseConnectSecure) { $body.UseConnectSecure } else { $false }
        $connectSecureBaseUrl = if ($body.ConnectSecureBaseUrl) { $body.ConnectSecureBaseUrl } else { Get-QueryParameter -Request $request -ParameterName "connectSecureBaseUrl" }
        $connectSecureTenant = if ($body.ConnectSecureTenant) { $body.ConnectSecureTenant } else { Get-QueryParameter -Request $request -ParameterName "connectSecureTenant" }
        $connectSecureClientId = if ($body.ConnectSecureClientId) { $body.ConnectSecureClientId } else { Get-QueryParameter -Request $request -ParameterName "connectSecureClientId" }
        $connectSecureClientSecret = if ($body.ConnectSecureClientSecret) { $body.ConnectSecureClientSecret } else { Get-QueryParameter -Request $request -ParameterName "connectSecureClientSecret" }
        $companyId = if ($body.CompanyId) { [int]$body.CompanyId } else { 0 }

        # Set defaults
        if ([string]::IsNullOrWhiteSpace($clientName)) {
            $clientName = "Client"
        }
        if ([string]::IsNullOrWhiteSpace($scanDate)) {
            $scanDate = (Get-Date).ToString("MM/dd/yyyy")
        }

        # Ensure temp directory exists
        if (-not (Test-Path $script:ApiConfig.TempDirectory)) {
            New-Item -Path $script:ApiConfig.TempDirectory -ItemType Directory -Force | Out-Null
        }

        # Determine data source
        $useConnectSecureAPI = $useConnectSecure -or (-not [string]::IsNullOrWhiteSpace($connectSecureBaseUrl))
        
        if ($useConnectSecureAPI) {
            # Validate ConnectSecure parameters
            if ([string]::IsNullOrWhiteSpace($connectSecureBaseUrl) -or 
                [string]::IsNullOrWhiteSpace($connectSecureTenant) -or 
                [string]::IsNullOrWhiteSpace($connectSecureClientId) -or 
                [string]::IsNullOrWhiteSpace($connectSecureClientSecret)) {
                Send-JsonResponse -Context $Context -Data @{ 
                    error = "ConnectSecure API parameters required: ConnectSecureBaseUrl, ConnectSecureTenant, ConnectSecureClientId, ConnectSecureClientSecret" 
                } -StatusCode 400
                return
            }

            # Connect to ConnectSecure API
            Write-ApiLog "Connecting to ConnectSecure API..." -Level Info
            $connected = Connect-ConnectSecureAPI -BaseUrl $connectSecureBaseUrl `
                                                    -TenantName $connectSecureTenant `
                                                    -ClientId $connectSecureClientId `
                                                    -ClientSecret $connectSecureClientSecret
            
            if (-not $connected) {
                Send-JsonResponse -Context $Context -Data @{ error = "Failed to authenticate with ConnectSecure API" } -StatusCode 401
                return
            }

            # Generate report from ConnectSecure data
            $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
            
            switch ($ReportType.ToLower()) {
                "pending-epss" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Pending EPSS Report_$timestamp.xlsx"
                    New-PendingEPSSReportFromConnectSecure -OutputPath $outputPath -CompanyId $companyId -ClientName $clientName
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "executive-summary" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Executive Summary_$timestamp.docx"
                    New-ExecutiveSummaryReportFromConnectSecure -OutputPath $outputPath -CompanyId $companyId -ClientName $clientName -ScanDate $scanDate
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
                "all-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName All Vulnerabilities Report_$timestamp.xlsx"
                    New-AllVulnerabilitiesReportFromConnectSecure -OutputPath $outputPath -CompanyId $companyId -ClientName $clientName
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "external-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName External Vulnerabilities Report_$timestamp.xlsx"
                    New-ExternalVulnerabilitiesReportFromConnectSecure -OutputPath $outputPath -CompanyId $companyId -ClientName $clientName
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "suppressed-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Suppressed Vulnerabilities Report_$timestamp.xlsx"
                    New-SuppressedVulnerabilitiesReportFromConnectSecure -OutputPath $outputPath -CompanyId $companyId -ClientName $clientName
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                default {
                    Send-JsonResponse -Context $Context -Data @{ error = "Unknown report type: $ReportType" } -StatusCode 400
                    return
                }
            }
        } else {
            # File-based report generation (existing logic)
            if ([string]::IsNullOrWhiteSpace($inputPath)) {
                Send-JsonResponse -Context $Context -Data @{ error = "InputPath parameter is required when not using ConnectSecure API" } -StatusCode 400
                return
            }

            if (-not (Test-Path $inputPath)) {
                Send-JsonResponse -Context $Context -Data @{ error = "Input file not found: $inputPath" } -StatusCode 404
                return
            }

            # Generate report based on type
            $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"

            switch ($ReportType.ToLower()) {
                "pending-epss" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Pending EPSS Report_$timestamp.xlsx"
                    New-ExcelReport -InputPath $inputPath -OutputPath $outputPath
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "executive-summary" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Executive Summary_$timestamp.docx"
                    New-ExecutiveSummaryReport -InputPath $inputPath -OutputPath $outputPath -ClientName $clientName -ScanDate $scanDate
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
                "all-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName All Vulnerabilities Report_$timestamp.xlsx"
                    New-AllVulnerabilitiesReport -InputPath $inputPath -OutputPath $outputPath
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "external-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName External Vulnerabilities Report_$timestamp.xlsx"
                    New-ExternalVulnerabilitiesReport -InputPath $inputPath -OutputPath $outputPath
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "suppressed-vulnerabilities" {
                    $outputPath = Join-Path $script:ApiConfig.TempDirectory "$clientName Suppressed Vulnerabilities Report_$timestamp.xlsx"
                    New-SuppressedVulnerabilitiesReport -InputPath $inputPath -OutputPath $outputPath
                    Send-FileResponse -Context $Context -FilePath $outputPath -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                default {
                    Send-JsonResponse -Context $Context -Data @{ error = "Unknown report type: $ReportType" } -StatusCode 400
                    return
                }
            }
        }

        # Clean up temp file after a delay (in background)
        Start-Job -ScriptBlock {
            param($FilePath)
            Start-Sleep -Seconds 60  # Wait 60 seconds before cleanup
            if (Test-Path $FilePath) {
                Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
            }
        } -ArgumentList $outputPath | Out-Null

    } catch {
        Write-ApiLog "Error handling report request: $($_.Exception.Message)" -Level Error
        Send-JsonResponse -Context $Context -Data @{ error = $_.Exception.Message } -StatusCode 500
    }
}

function Handle-HealthCheck {
    param(
        [System.Net.HttpListenerContext]$Context
    )

    Send-JsonResponse -Context $Context -Data @{
        status = "healthy"
        version = "1.0.0"
        timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
}

# --- Main API Server ---

function Start-ApiServer {
    param(
        [int]$Port = 8080,
        [string]$BindHost = "localhost"
    )

    $listener = New-Object System.Net.HttpListener
    $prefix = "http://${BindHost}:${Port}/"
    $listener.Prefixes.Add($prefix)

    try {
        $listener.Start()
        Write-ApiLog "VScanMagic API Server started on $prefix" -Level Success
        Write-ApiLog "Available endpoints:" -Level Info
        Write-ApiLog "  GET  /health - Health check" -Level Info
        Write-ApiLog "  POST /api/reports/pending-epss - Generate Pending EPSS Report (XLSX)" -Level Info
        Write-ApiLog "  POST /api/reports/executive-summary - Generate Executive Summary (DOCX)" -Level Info
        Write-ApiLog "  POST /api/reports/all-vulnerabilities - Generate All Vulnerabilities Report (XLSX)" -Level Info
        Write-ApiLog "  POST /api/reports/external-vulnerabilities - Generate External Vulnerabilities Report (XLSX)" -Level Info
        Write-ApiLog "  POST /api/reports/suppressed-vulnerabilities - Generate Suppressed Vulnerabilities Report (XLSX)" -Level Info

        while ($listener.IsListening) {
            try {
                $context = $listener.GetContext()
                $request = $context.Request
                $path = $request.Url.AbsolutePath
                $method = $request.HttpMethod

                Write-ApiLog "$method $path" -Level Info

                # Handle CORS preflight
                if ($method -eq "OPTIONS") {
                    $context.Response.Headers.Add("Access-Control-Allow-Origin", "*")
                    $context.Response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                    $context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
                    $context.Response.StatusCode = 200
                    $context.Response.OutputStream.Close()
                    continue
                }

                # Add CORS headers
                $context.Response.Headers.Add("Access-Control-Allow-Origin", "*")

                # Route requests
                if ($path -eq "/health" -and $method -eq "GET") {
                    Handle-HealthCheck -Context $context
                }
                elseif ($path -match "^/api/reports/(.+)$") {
                    $reportType = $matches[1]
                    Handle-ReportRequest -Context $context -ReportType $reportType
                }
                else {
                    Send-JsonResponse -Context $context -Data @{ error = "Not found" } -StatusCode 404
                }

            } catch {
                Write-ApiLog "Error processing request: $($_.Exception.Message)" -Level Error
                if ($context) {
                    try {
                        Send-JsonResponse -Context $context -Data @{ error = "Internal server error" } -StatusCode 500
                    } catch { }
                }
            }
        }
    } catch {
        Write-ApiLog "Failed to start API server: $($_.Exception.Message)" -Level Error
    } finally {
        if ($listener.IsListening) {
            $listener.Stop()
        }
        Write-ApiLog "API Server stopped" -Level Info
    }
}

# --- Start Server ---
Write-ApiLog "Starting VScanMagic API Server..." -Level Info
Start-ApiServer -Port $script:ApiConfig.Port -BindHost $script:ApiConfig.Host
