# ConnectSecure-API.Part3.ps1 — Dot-sourced by ConnectSecure-API.ps1 (do not run directly)
# Report generation from ConnectSecure, report builder jobs, batch download, local fallback.
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
    param(
        [string]$OutputPath,
        [int]$CompanyId = 0,
        [string]$ClientName = 'Client',
        [array]$VulnerabilityData = $null,
        [scriptblock]$OnProgress = $null,
        [int]$DebugLimit = 0,
        [switch]$IncludeRegistryAndNetwork = $true
    )
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    # Fetch assets once and share: asset_wise uses for company filtering, enrichment uses for OS/Owner/asset names
    $assets = @()
    try {
        $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
    } catch {
        Write-CSApiLog ('Asset fetch failed (OS/Owner will be empty): ' + $_.Exception.Message) -Level Warning
    }
    # Use asset_wise_vulnerabilities for 13-column format (CVE ID, Severity, CVSS, EPSS, Asset Name, OS, IP, Description, App Name, Product Name, First Seen, Last Seen, Owner, Fix)
    $assetWise = Get-ConnectSecureAssetWiseVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -Assets $assets
    if ($null -eq $assetWise) { $assetWise = @() }
    $assetDetailMap = Get-ConnectSecureAssetDetailMap -Assets $assets
    $allVulns = ConvertTo-ConnectSecure13ColumnFormat -AssetWiseData $assetWise -AssetDetailMap $assetDetailMap
    if ($null -eq $allVulns) { $allVulns = @() }

    # Merge Registry (open) and Network into All Vulnerabilities for Top N / ticket instructions
    if ($IncludeRegistryAndNetwork) {
        $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        try {
            $registryVulns = Get-ConnectSecureRegistryVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll
            if ($registryVulns -and $registryVulns.Count -gt 0) {
                $registry13 = ConvertFrom-RegistryProblemsFormat -Data $registryVulns -AssetDetailMap $assetDetailMap -AssetMap $assetMap
                if ($registry13 -and $registry13.Count -gt 0) {
                    $allVulns = @($allVulns) + @($registry13)
                    Write-CSApiLog "Merged $($registry13.Count) registry vulnerabilities into All Vulnerabilities" -Level Info
                }
            }
        } catch {
            Write-CSApiLog ('Registry merge skipped: ' + $_.Exception.Message) -Level Warning
        }
        try {
            $networkVulns = Get-ConnectSecureNetworkVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll
            if ($networkVulns -and $networkVulns.Count -gt 0) {
                $network13 = ConvertFrom-NetworkTo13ColumnFormat -Data $networkVulns
                if ($network13 -and $network13.Count -gt 0) {
                    $allVulns = @($allVulns) + @($network13)
                    Write-CSApiLog "Merged $($network13.Count) network vulnerabilities into All Vulnerabilities" -Level Info
                }
            }
        } catch {
            Write-CSApiLog ('Network merge skipped: ' + $_.Exception.Message) -Level Warning
        }
    }

    if ($allVulns.Count -eq 0) {
        # Empty export needs All Vulnerabilities 13-column headers for Get-VulnerabilityData compatibility
        $emptyRow = [PSCustomObject]@{
            'Source'=''; 'CVE ID'=''; 'Severity'=''; 'CVSS Score'=''; 'EPSS Score'=''; 'Asset Name'=''; 'OS'=''; 'IP Address'=''; 'Description'=''; 'Application Name'=''; 'Product Name'=''; 'First Seen'=''; 'Last Seen'=''; 'Owner'=''; 'Fix'=''
        }
        Export-ConnectSecureDataToExcel -Data @($emptyRow) -OutputPath $OutputPath -SheetName 'All Vulnerabilities' -OnProgress $null
        return
    }
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

function New-RegistryVulnerabilitiesReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $registryVulns = Get-ConnectSecureRegistryVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll
    if ($null -eq $registryVulns) { $registryVulns = @() }
    $assetDetailMap = @{}; $assetMap = @{}
    if ($registryVulns.Count -gt 0) {
        try {
            $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
            $assetDetailMap = Get-ConnectSecureAssetDetailMap -Assets $assets
            $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        } catch { Write-CSApiLog ('Asset maps skipped: ' + $_.Exception.Message) -Level Warning }
    }
    $converted = ConvertFrom-RegistryProblemsFormat -Data $registryVulns -AssetDetailMap $assetDetailMap -AssetMap $assetMap
    Export-ConnectSecureDataToExcel -Data $converted -OutputPath $OutputPath -SheetName 'Registry Vulnerabilities' -OnProgress $null
}

function New-RegistryRemediatedReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $registryVulns = Get-ConnectSecureRegistryVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -RemediatedOnly
    if ($null -eq $registryVulns) { $registryVulns = @() }
    $assetDetailMap = @{}; $assetMap = @{}
    if ($registryVulns.Count -gt 0) {
        try {
            $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
            $assetDetailMap = Get-ConnectSecureAssetDetailMap -Assets $assets
            $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        } catch { Write-CSApiLog ('Asset maps skipped: ' + $_.Exception.Message) -Level Warning }
    }
    $converted = ConvertFrom-RegistryProblemsFormat -Data $registryVulns -AssetDetailMap $assetDetailMap -AssetMap $assetMap
    Export-ConnectSecureDataToExcel -Data $converted -OutputPath $OutputPath -SheetName 'Registry Remediated' -OnProgress $null
}

function New-RegistrySuppressedReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $registryVulns = Get-ConnectSecureRegistryVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll -SuppressedOnly
    if ($null -eq $registryVulns) { $registryVulns = @() }
    $assetDetailMap = @{}; $assetMap = @{}
    if ($registryVulns.Count -gt 0) {
        try {
            $assets = Get-ConnectSecureAssets -CompanyId $CompanyId -Limit 5000 -FetchAll:$true
            $assetDetailMap = Get-ConnectSecureAssetDetailMap -Assets $assets
            $assetMap = Get-ConnectSecureAssetIdToNameMap -Assets $assets
        } catch { Write-CSApiLog ('Asset maps skipped: ' + $_.Exception.Message) -Level Warning }
    }
    $converted = ConvertFrom-RegistryProblemsFormat -Data $registryVulns -AssetDetailMap $assetDetailMap -AssetMap $assetMap
    Export-ConnectSecureDataToExcel -Data $converted -OutputPath $OutputPath -SheetName 'Registry Suppressed' -OnProgress $null
}

function New-NetworkVulnerabilitiesReportFromConnectSecure {
    param([string]$OutputPath, [int]$CompanyId = 0, [string]$ClientName = 'Client', [scriptblock]$OnProgress = $null, [int]$DebugLimit = 0)
    $limit = if ($DebugLimit -gt 0) { $DebugLimit } else { 5000 }
    $fetchAll = ($DebugLimit -le 0)
    $networkVulns = Get-ConnectSecureNetworkVulnerabilities -CompanyId $CompanyId -Limit $limit -FetchAll:$fetchAll
    if ($null -eq $networkVulns) { $networkVulns = @() }
    $converted = ConvertFrom-NetworkVulnerabilitiesFormat -Data $networkVulns
    Export-ConnectSecureDataToExcel -Data $converted -OutputPath $OutputPath -SheetName 'Network Vulnerabilities' -OnProgress $null
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
    Tries both isGlobal=true and isGlobal=false, merges results.
    Response: varies; uses recursive extraction for id+reportType objects.
    #>
    param(
        [bool]$IsGlobal = $false,
        [int]$CompanyId = 0,
        [int]$Skip = 0,
        [int]$Limit = 2000,
        [switch]$UseGlobalOnly = $false
    )
    if ($CompanyId -ne 0) { $IsGlobal = $false }

    $toTry = @()
    if ($UseGlobalOnly) { $toTry = @($true) }
    else { $toTry = @($IsGlobal, (-not $IsGlobal)) }

    $collected = [System.Collections.ArrayList]@()
    $seenIds = @{}

    foreach ($useGlobal in $toTry) {
        $queryParams = @{ isGlobal = $useGlobal.ToString().ToLower(); skip = $Skip; limit = $Limit }
        try {
            $response = Invoke-ConnectSecureRequest -Endpoint '/report_builder/standard_reports' -Method GET -QueryParameters $queryParams
            if (-not $response.status) { continue }
            $msg = $response.message
            if ($msg) {
                $sections = @($msg)
                foreach ($sec in $sections) {
                    if (-not $sec) { continue }
                    $reportCategories = $null
                    if ($sec.Reports) { $reportCategories = $sec.Reports }
                    if (-not $reportCategories) { continue }
                    foreach ($cat in $reportCategories) {
                        $desc = 'Unknown'
                        if ($cat.description) { $desc = $cat.description.ToString() }
                        $descLower = $desc.ToLower()
                        $reportsList = $cat.reports
                        if (-not $reportsList) { continue }
                        foreach ($rep in $reportsList) {
                            $rid = $rep.id
                            if (-not $rid -and $rep.reportId) { $rid = $rep.reportId }
                            $rt = $rep.reportType
                            if (-not $rt -and $rep.report_type) { $rt = $rep.report_type }
                            if (-not $rid -or -not $rt) { continue }
                            $rtStr = $rt.ToString().ToLower()
                            if ($rtStr -ne 'xlsx' -and $rtStr -ne 'docx' -and $rtStr -ne 'pdf') { continue }
                            $key = "$rid-$rtStr"
                            if ($seenIds[$key]) { continue }
                            $seenIds[$key] = $true
                            $rep | Add-Member -NotePropertyName '_category' -NotePropertyValue $descLower -Force
                            $rep | Add-Member -NotePropertyName '_categoryDisplay' -NotePropertyValue $desc -Force
                            [void]$collected.Add($rep)
                        }
                    }
                }
            }
            if ($collected.Count -eq 0 -and $response.data) {
                foreach ($d in @($response.data)) {
                    $dDesc = 'Unknown'
                    if ($d.description) { $dDesc = $d.description }
                    $rid = $d.id; $rt = $d.reportType
                    if ($rid -and $rt) {
                        $rtStr = $rt.ToString().ToLower()
                        if ($rtStr -eq 'xlsx' -or $rtStr -eq 'docx' -or $rtStr -eq 'pdf') {
                            $d | Add-Member -NotePropertyName '_categoryDisplay' -NotePropertyValue $dDesc -Force
                            [void]$collected.Add($d)
                        }
                    }
                }
            }
            if ($collected.Count -gt 0) { break }
        } catch {
            Write-CSApiLog ('standard_reports isGlobal=' + $useGlobal + ': ' + $_.Exception.Message) -Level Warning
        }
    }
    if ($collected.Count -gt 0) {
        Write-CSApiLog ('Standard reports: ' + $collected.Count + ' items') -Level Info
    }
    return @($collected)
}

function Get-StandardReportIdForType {
    <#
    .SYNOPSIS
    Resolves our InternalReportType to a report ID from standard_reports API.
    Uses only tenant-provided standard reports - no hardcoded IDs.
    #>
    param(
        [string]$InternalReportType,
        [string]$ReportFormat = 'xlsx',
        [int]$CompanyId = 0
    )
    $reports = Get-ConnectSecureStandardReports -CompanyId $CompanyId
    if (-not $reports -or $reports.Count -eq 0) {
        Write-CSApiLog 'No standard reports returned - cannot create report' -Level Warning
        return $null
    }
    $wantFormat = if ($ReportFormat -eq 'docx') { 'docx' } elseif ($ReportFormat -eq 'pdf') { 'pdf' } else { 'xlsx' }
    $categoryPatterns = @{
        'all-vulnerabilities' = 'all vulnerabilities report'
        'suppressed-vulnerabilities' = 'suppressed vulnerabilities'
        'external-vulnerabilities' = 'external scan'
        'executive-summary' = 'executive summary report'
        'pending-epss' = 'pending remediation epss score reports'
        'network-vulnerabilities' = 'network scan findings'
    }
    $pattern = $categoryPatterns[$InternalReportType]
    $candidates = @()
    foreach ($r in $reports) {
        $rt = if ($r.reportType) { $r.reportType.ToString().ToLower() } else { '' }
        if ($rt -ne $wantFormat) { continue }
        $cat = if ($r._category) { $r._category.ToString().ToLower() } else { '' }
        if ($pattern -and $cat -like ('*' + $pattern + '*')) {
            $candidates += $r
            break
        }
    }
    if ($candidates.Count -eq 0 -and $pattern) {
        foreach ($r in $reports) {
            $rt = if ($r.reportType) { $r.reportType.ToString().ToLower() } else { '' }
            if ($rt -ne $wantFormat) { continue }
            $cat = if ($r._category) { $r._category.ToString().ToLower() } else { '' }
            if ($cat -match ($pattern -replace '\s+', '.*')) {
                $candidates += $r
                break
            }
        }
    }
    if ($candidates.Count -eq 0 -and $pattern) {
        foreach ($r in $reports) {
            $rt = if ($r.reportType) { $r.reportType.ToString().ToLower() } else { '' }
            if ($rt -ne $wantFormat) { continue }
            $disp = ''
            foreach ($p in @('displayReportName','displayName','description','name')) { if ($r.$p) { $disp = [string]$r.$p; break } }
            $disp = $disp.ToLower()
            if ($disp -like ('*' + $pattern + '*')) {
                $candidates += $r
                break
            }
        }
    }
    if ($candidates.Count -eq 0) {
        $formatMatches = @($reports | Where-Object { ($_.reportType -or '').ToString().ToLower() -eq $wantFormat })
        if ($formatMatches.Count -eq 1) {
            $id = $formatMatches[0].id
            if ($id) { Write-CSApiLog ('Using only ' + $wantFormat + ' report from standard_reports: ' + $id) -Level Info; return $id }
        }
        if ($formatMatches.Count -gt 0) {
            $id = $formatMatches[0].id
            if ($id) { Write-CSApiLog ('Using first ' + $wantFormat + ' report (no name match for ' + $InternalReportType + '): ' + $id) -Level Info; return $id }
        }
        # Fallback: use known IDs when dynamic match fails (API structure may vary by tenant)
        if ($categoryPatterns[$InternalReportType]) {
            $knownIds = @{
                'all-vulnerabilities'        = '00000000000000000000000000000000'
                'suppressed-vulnerabilities'   = '1d091564830b44c485a0ddc35ace9ac6'
                'external-vulnerabilities'    = '01beb6b930744e11b690bb9dc25118fb'
                'executive-summary'           = '1cd4f45884264d15bee4173dc58b6a57'
                'pending-epss'                = '85d4913c0dbc4fc782b858f0d27dd180'
            }
            $knownId = $knownIds[$InternalReportType]
            $exists = $reports | Where-Object { $_.id -eq $knownId }
            if ($knownId -and $exists) {
                Write-CSApiLog ('Using known report id for ' + $InternalReportType + ': ' + $knownId) -Level Info
                return $knownId
            }
        }
        Write-CSApiLog ('No ' + $wantFormat + ' report found in standard_reports for ' + $InternalReportType) -Level Warning
        return $null
    }
    $id = $candidates[0].id
    if ($id) { Write-CSApiLog ('Matched ' + $InternalReportType + ' to standard report: ' + $id) -Level Info }
    return $id
}

function New-ConnectSecureReportJob {
    <#
    .SYNOPSIS
    Creates a report generation job in ConnectSecure.
    Tries multiple endpoints and parameter formats for compatibility.
    Swagger: POST /report_builder/create_report_job
    .PARAMETER ReportType
    Report id (string) or reportType from standard_reports.
    .PARAMETER CompanyId
    Company ID (0 for all companies).
    .PARAMETER ReportFormat
    Format: xlsx, docx, pdf.
    .PARAMETER ReportName
    Optional display name (reportName in camelCase format).
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$ReportType,
        [int]$CompanyId = 0,
        [string]$ReportFormat = 'xlsx',
        [string]$ReportName = '',
        [string]$ClientName = '',
        [hashtable]$AdditionalParams = @{}
    )

    $reportIdStr = $null
    if ($ReportType -is [int]) {
        $reportIdStr = $ReportType.ToString()
    } elseif ($ReportType -match '^\d+$') {
        $reportIdStr = $ReportType.ToString()
    } elseif ($ReportType -match '^[a-fA-F0-9]{32}$') {
        $reportIdStr = $ReportType.ToString()
    } elseif ($ReportType) {
        $reportIdStr = $ReportType.ToString()
    }

    # Company context: portal uses company_id="global" (string) for global reports, numeric ID for company reports
    $companyIdParam = $CompanyId
    $companyName = "Company $CompanyId"
    if ($CompanyId -eq 0) {
        $companyIdParam = 'global'
        $companyName = 'Global'
    } elseif (-not [string]::IsNullOrWhiteSpace($ClientName) -and $ClientName -ne "Company" -and $ClientName -ne "All Companies") {
        $companyName = $ClientName.Trim()
    } else {
        try {
            $compResp = Invoke-ConnectSecureRequest -Endpoint '/r/company/companies' -Method GET -QueryParameters @{ limit = 5000; skip = 0 }
            $companies = @()
            if ($compResp.data) { $companies = @($compResp.data) }
            $match = $companies | Where-Object { $cid = $_.id -or $_.company_id -or $_.companyId; [int]$cid -eq $CompanyId } | Select-Object -First 1
            if ($match) {
                $companyName = ($match.name -or $match.company_name -or $match.companyName -or '').ToString().Trim()
                if ([string]::IsNullOrWhiteSpace($companyName)) { $companyName = "Company $CompanyId" }
            }
        } catch { }
    }

    # Portal-style reportName: no spaces (e.g. AllVulnerabilitiesReport)
    $reportNameCompact = ($ReportName -replace '\s', '')
    if ([string]::IsNullOrWhiteSpace($reportNameCompact)) { $reportNameCompact = 'Report' }

    # Build body variants - portal format FIRST (reportType=Standard, isFilter=true)
    # Per portal capture: company_id is "global" (string) for global reports, numeric for company reports
    $bodyPortal = @{
        reportId     = $reportIdStr
        reportName   = $reportNameCompact
        reportType   = 'Standard'
        isFilter     = $true
        fileType     = $ReportFormat
        reportFilter = @{}
        company_id   = $companyIdParam
        company_name = $companyName
    }
    foreach ($k in $AdditionalParams.Keys) { $bodyPortal[$k] = $AdditionalParams[$k] }

    $bodySnake = @{ company_id = $companyIdParam; report_format = $ReportFormat }
    if ($reportIdStr -match '^[a-fA-F0-9]{32}$') {
        $bodySnake.report_id = $reportIdStr
    } else {
        $bodySnake.reportType = $reportIdStr
    }
    $bodySnake += $AdditionalParams

    $bodyCamel = @{
        company_id = $companyIdParam
        reportId = $reportIdStr
        reportType = $ReportFormat
        fileType = $ReportFormat
    }
    if (-not [string]::IsNullOrWhiteSpace($ReportName)) { $bodyCamel.reportName = $ReportName }
    foreach ($k in $AdditionalParams.Keys) { $bodyCamel[$k] = $AdditionalParams[$k] }

    $endpoints = @('/report_builder/create_report_job', '/r/report_builder/create_report_job')
    $bodyVariants = @($bodyPortal, $bodySnake, $bodyCamel)
    $lastErr = $null

    foreach ($ep in $endpoints) {
        foreach ($body in $bodyVariants) {
            try {
                $response = Invoke-ConnectSecureRequest -Endpoint $ep -Method POST -Body $body
                if ($response.status -eq $false -or (-not $response.status -and $null -ne $response.status)) {
                    $msg = if ($response.message) { $response.message } else { 'status=false' }
                    $msgStr = if ($msg -is [string]) { $msg } else { $msg | ConvertTo-Json -Compress }
                    $lastErr = $msgStr
                    if ($msgStr -match 'Please Contact Support') { continue }
                    continue  # try next variant
                }

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
                        if ($msg -is [string] -and $msg -match '^[a-fA-F0-9-]{36}$') { $jobId = $msg }
                        elseif ($msg -is [string] -and $msg -match '^\d+$') { $jobId = $msg }
                        elseif ($null -ne $msg.job_id) { $jobId = $msg.job_id }
                    }
                } catch { }

                if (-not [string]::IsNullOrWhiteSpace($jobId)) {
                    Write-CSApiLog ('Report job created: ' + $jobId + ' (company_id=' + $CompanyId + ')') -Level Success
                    return $jobId.ToString()
                }
            } catch {
                $lastErr = $_.Exception.Message
                if ($lastErr -match 'Please Contact Support') { continue }
                throw $_
            }
        }
    }

    if ($lastErr -match 'Please Contact Support') {
        Write-CSApiLog 'Report Builder creation may be disabled for API. Use Download by Job ID with a report created in the portal.' -Level Warning
    }
    throw 'ConnectSecure create_report_job failed. ' + (if ($lastErr) { $lastErr } else { 'Try Download by Job ID for reports created in the portal.' })
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
    $jobIdArray = '["' + $JobId + '"]'  # Portal sends job_id as JSON array
    $paramVariants = @(
        @{ job_id = $jobIdArray; isGlobal = $isGlob },
        @{ job_id = $JobId; isGlobal = $isGlob }
    )
    if (-not $IsGlobal -and $CompanyId -ne 0) {
        $paramVariants += @{ job_id = $jobIdArray; isGlobal = $isGlob; company_id = $CompanyId }
        $paramVariants += @{ job_id = $JobId; isGlobal = $isGlob; company_id = $CompanyId }
    }
    $endpoints = @('/report_builder/get_report_link', '/r/report_builder/get_report_link')
    $response = $null
    $lastErr = $null
    foreach ($ep in $endpoints) {
        foreach ($qp in $paramVariants) {
            try {
                # RetryCount 1 - our polling loop handles retries with delay; 404 means report still generating
                $response = Invoke-ConnectSecureRequest -Endpoint $ep -Method GET -QueryParameters $qp -RetryCount 1
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

function Invoke-ConnectSecureReportDownloadByJobId {
    <#
    .SYNOPSIS
    Downloads a ConnectSecure report by job ID using get_report_link.
    Use when report creation via API fails (e.g. Please Contact Support) but you have a job ID from the portal.
    Pre-signed R2/S3 URLs are downloaded without Authorization header (prevents 400 Bad Request).
    .PARAMETER JobId
    Report job ID (GUID format, e.g. from ConnectSecure portal).
    .PARAMETER CompanyId
    Company ID (0 for global/all companies).
    .PARAMETER OutputPath
    Full path for the downloaded file. If not provided, saves to OutputDir with auto-generated filename.
    .PARAMETER OutputDir
    Directory to save the file when OutputPath is not specified.
    .OUTPUTS
    Full path of the downloaded file.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$JobId,
        [int]$CompanyId = 0,
        [string]$OutputPath = '',
        [string]$OutputDir = '.'
    )
    $JobId = $JobId.Trim()
    if ([string]::IsNullOrWhiteSpace($JobId)) { throw 'JobId is required' }

    $downloadUrl = Get-ConnectSecureReportLink -JobId $JobId -IsGlobal ($CompanyId -eq 0) -CompanyId $CompanyId
    if (-not $downloadUrl.StartsWith('http')) {
        $slash = [char]47
        $base = $script:ConnectSecureConfig.BaseUrl.TrimEnd($slash)
        $path = $downloadUrl.TrimStart($slash)
        $downloadUrl = $base + '/' + $path
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $ext = if ($downloadUrl -match '\.(xlsx|xls|docx|doc|pdf|zip)(?:\?|$)') { $Matches[1] } else { 'xlsx' }
        $safeDir = [System.IO.Path]::GetFullPath($OutputDir)
        if (-not (Test-Path $safeDir)) { New-Item -ItemType Directory -Path $safeDir -Force | Out-Null }
        $ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $fileName = 'Report-' + $JobId + '-' + $ts + '.' + $ext
        $OutputPath = Join-Path $safeDir $fileName
    }

    # Pre-signed R2/S3 URLs authenticate via query params - do NOT send Authorization (causes 400)
    $isPresigned = $downloadUrl -match 'r2\.cloudflarestorage|X-Amz-Signature'
    $headers = @{}
    if (-not $isPresigned) {
        $headers['Authorization'] = 'Bearer ' + $script:ConnectSecureConfig.AccessToken
        if (-not [string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.UserId)) {
            $headers['X-USER-ID'] = $script:ConnectSecureConfig.UserId.ToString()
        }
    }
    Invoke-WebRequest -Uri $downloadUrl -Method GET -Headers $headers -OutFile $OutputPath -UseBasicParsing
    Write-CSApiLog ('Downloaded report by Job ID to ' + $OutputPath) -Level Success
    return $OutputPath
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
        [double]$MinEPSS = 0,
        [bool]$IncludeCritical = $true,
        [bool]$IncludeHigh = $true,
        [bool]$IncludeMedium = $true,
        [bool]$IncludeLow = $true,
        [scriptblock]$OnProgress = $null,
        [int]$DebugLimit = 0,
        [switch]$SkipPostDownloadTopX = $false
    )

    function Update-Prog { param($m) if ($OnProgress) { & $OnProgress $m } }

    $standardTypes = @('all-vulnerabilities', 'suppressed-vulnerabilities', 'external-vulnerabilities', 'executive-summary', 'pending-epss', 'registry-vulnerabilities', 'registry-remediated', 'registry-suppressed', 'network-vulnerabilities')
    $localOnlyTypes = @('registry-vulnerabilities', 'registry-remediated', 'registry-suppressed')
    $reportsWithType = $Reports | Where-Object { $_.Type -and -not $_.ReportId }
    foreach ($r in $reportsWithType) {
        if ($r.Type -notin $standardTypes) {
            throw ('Unknown report type: ' + $r.Type + '. Standard reports: all-vulnerabilities, suppressed-vulnerabilities, external-vulnerabilities, executive-summary, pending-epss, registry-vulnerabilities, registry-remediated, registry-suppressed, network-vulnerabilities')
        }
    }

    $script:CSReportBuilderUnavailable = $false
    $script:LastReportJobId = $null

    $succeeded = [System.Collections.ArrayList]::new()
    $failed = [System.Collections.ArrayList]::new()
    $companyLabel = if ($CompanyId -eq 0) { 'All Companies' } else { ('company ' + $CompanyId) }

    # Generate registry and network locally (no Report Builder support)
    foreach ($report in $Reports) {
        if ($report.Type -in $localOnlyTypes) {
            $path = & $OutputPathTemplate $report
            try {
                Update-Prog ('Generating ' + $report.Name + ' from API...')
                Invoke-LocalReportFallback -InternalReportType $report.Type -OutputPath $path -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate -OnProgress $OnProgress -DebugLimit $DebugLimit
                $null = $succeeded.Add($report)
                Write-CSApiLog ('Generated: ' + $path) -Level Success
            } catch {
                $errText = if ($null -eq $_.Exception.Message) { 'Unknown error' } elseif ($_.Exception.Message -is [array]) { ($_.Exception.Message -join '; ') } else { [string]$_.Exception.Message }
                Write-CSApiLog ('Local report failed (' + $report.Name + '): ' + $errText) -Level Error
                $null = $failed.Add([PSCustomObject]@{ Report = $report; Path = $path; Error = $errText })
            }
        }
    }
    $reportBuilderReports = @($Reports | Where-Object { $_.Type -notin $localOnlyTypes })
    if ($reportBuilderReports.Count -eq 0) {
        return @{ Succeeded = [array]$succeeded; Failed = [array]$failed }
    }
    $isGlobal = ($CompanyId -eq 0)
    $pollInterval = 2
    $maxWaitSeconds = 600

    # Phase 1: Create all report jobs in parallel (all start generating on server immediately)
    # Skip local-only types (registry-*, network-vulnerabilities) - they use Invoke-LocalReportFallback above
    $pending = [System.Collections.ArrayList]::new()
    Update-Prog ('Creating ' + $reportBuilderReports.Count + ' report jobs for ' + $companyLabel + '...')
    foreach ($report in $reportBuilderReports) {
        if ($script:CSReportBuilderUnavailable) { break }
        $path = & $OutputPathTemplate $report
        try {
            $reportFormat = if ($report.Ext -eq 'docx') { 'docx' } elseif ($report.Ext -eq 'pdf') { 'pdf' } else { 'xlsx' }
            $ext = $reportFormat
            $reportId = $null
            $reportName = $report.Name
            if ($report.ReportId) {
                $reportId = $report.ReportId
            } else {
                $reportId = Get-StandardReportIdForType -InternalReportType $report.Type -ReportFormat $ext -CompanyId $CompanyId
                if (-not $reportId) {
                    throw ('No standard report match for ' + $report.Type)
                }
                $reportName = $script:CSReportNameMap[$report.Type]
            }
            $jobId = New-ConnectSecureReportJob -ReportType $reportId -CompanyId $CompanyId -ReportFormat $ext -ReportName $reportName -ClientName $ClientName
            $script:LastReportJobId = $jobId
            $null = $pending.Add([PSCustomObject]@{ Report = $report; JobId = $jobId; Path = $path })
        } catch {
            $errText = if ($null -eq $_.Exception.Message) { 'Unknown error' } elseif ($_.Exception.Message -is [array]) { ($_.Exception.Message -join '; ') } else { [string]$_.Exception.Message }
            if ($errText -match 'Please Contact Support') { $script:CSReportBuilderUnavailable = $true }
            Write-CSApiLog ('Report job creation failed (' + $report.Name + '): ' + $errText) -Level Error
            $null = $failed.Add([PSCustomObject]@{ Report = $report; Path = $path; Error = $errText })
        }
    }

    # Phase 2: Poll all jobs and download as each becomes ready
    if ($pending.Count -gt 0) {
        Start-Sleep -Seconds 5
        $start = Get-Date
        while ($pending.Count -gt 0) {
            $elapsed = ((Get-Date) - $start).TotalSeconds
            if ($elapsed -ge $maxWaitSeconds) {
                foreach ($p in $pending) {
                    Write-CSApiLog ('Report timed out (' + $p.Report.Name + '). Job ID ' + $p.JobId + ' - use Download by Job ID when ready.') -Level Warning
                    $null = $failed.Add([PSCustomObject]@{ Report = $p.Report; Path = $p.Path; Error = 'Timed out' })
                }
                $pending.Clear()
                break
            }
            $stillPending = [System.Collections.ArrayList]::new()
            foreach ($p in $pending) {
                Update-Prog ('Waiting for reports... (' + [int]$elapsed + 's, ' + $pending.Count + ' pending)')
                $downloadUrl = $null
                try {
                    $downloadUrl = Get-ConnectSecureReportLink -JobId $p.JobId -IsGlobal $isGlobal -CompanyId $CompanyId
                } catch {
                    if ($_.Exception.Message -match '404') { $null = $stillPending.Add($p); continue }
                    $null = $failed.Add([PSCustomObject]@{ Report = $p.Report; Path = $p.Path; Error = $_.Exception.Message })
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($downloadUrl)) { $null = $stillPending.Add($p); continue }
                if (-not $downloadUrl.StartsWith('http')) {
                    $base = $script:ConnectSecureConfig.BaseUrl.TrimEnd('/')
                    $downloadUrl = $base + '/' + $downloadUrl.TrimStart('/')
                }
                $isPresigned = $downloadUrl -match 'r2\.cloudflarestorage|X-Amz-Signature'
                $headers = @{}
                if (-not $isPresigned) {
                    $headers['Authorization'] = 'Bearer ' + $script:ConnectSecureConfig.AccessToken
                    if (-not [string]::IsNullOrWhiteSpace($script:ConnectSecureConfig.UserId)) {
                        $headers['X-USER-ID'] = $script:ConnectSecureConfig.UserId.ToString()
                    }
                }
                try {
                    Invoke-WebRequest -Uri $downloadUrl -Method GET -Headers $headers -OutFile $p.Path -UseBasicParsing
                    Write-CSApiLog ('Downloaded: ' + $p.Path) -Level Success
                    $null = $succeeded.Add($p.Report)
                } catch {
                    if ($_.Exception.Message -match '404') {
                        $null = $stillPending.Add($p)
                    } else {
                        $null = $failed.Add([PSCustomObject]@{ Report = $p.Report; Path = $p.Path; Error = $_.Exception.Message })
                    }
                }
            }
            $pending = $stillPending
            if ($pending.Count -gt 0) { Start-Sleep -Seconds $pollInterval }
        }
    }

    # Post-download: generate Top X report from vulnerability XLSX (when available)
    # Skip when caller will process and generate their own (e.g. GUI download+process flow)
    if (-not $SkipPostDownloadTopX) {
    $vulnReport = $succeeded | Where-Object { $_.Type -eq 'all-vulnerabilities' } | Select-Object -First 1
    if (-not $vulnReport) {
        $vulnReport = $succeeded | Where-Object { $_.Type -in @('external-vulnerabilities', 'suppressed-vulnerabilities') } | Select-Object -First 1
    }
    if ($null -ne $vulnReport -and $vulnReport.Ext -eq 'xlsx') {
        $avPath = & $OutputPathTemplate $vulnReport
        if (Test-Path -LiteralPath $avPath) {
            $topTitle = if ($TopCount -le 0) { 'Top Vulnerabilities Report' } elseif ($TopCount -eq 10) { 'Top Ten Vulnerabilities Report' } else { "Top $TopCount Vulnerabilities Report" }
            Update-Prog ("Generating $topTitle from $($vulnReport.Name)...")
            try {
                if ((Get-Command -Name 'Get-VulnerabilityData' -ErrorAction SilentlyContinue) -and
                    (Get-Command -Name 'Get-Top10Vulnerabilities' -ErrorAction SilentlyContinue) -and
                    (Get-Command -Name 'New-WordReport' -ErrorAction SilentlyContinue)) {
                    $vulnData = Get-VulnerabilityData -ExcelPath $avPath
                    if ($null -ne $vulnData -and $vulnData.Count -gt 0) {
                        $top10 = Get-Top10Vulnerabilities -VulnData $vulnData -Count $TopCount `
                            -MinEPSS $MinEPSS -IncludeCritical $IncludeCritical -IncludeHigh $IncludeHigh `
                            -IncludeMedium $IncludeMedium -IncludeLow $IncludeLow
                        $outputDir = Split-Path -Path $avPath -Parent
                        $stem = [System.IO.Path]::GetFileNameWithoutExtension($avPath)
                        $reportNamePart = [regex]::Escape($vulnReport.Name)
                        $topXStem = $stem -replace (' - ' + $reportNamePart + ' - '), (" - $topTitle - ")
                        $topXPath = Join-Path $outputDir ($topXStem + '.docx')
                        New-WordReport -OutputPath $topXPath -ClientName $ClientName -ScanDate $ScanDate -Top10Data $top10 -TimeEstimates $null -IsRMITPlus $false -GeneralRecommendations @() -ReportTitle $topTitle
                        Write-CSApiLog ('Generated: ' + $topXPath) -Level Success
                        $null = $succeeded.Add([PSCustomObject]@{ Type = 'top-vulnerabilities'; Name = $topTitle; Ext = 'docx' })
                    } else {
                        Write-CSApiLog ('No vulnerability data found in ' + $avPath + ' - skipping Top X generation') -Level Warning
                    }
                } else {
                    Write-CSApiLog 'Report generation functions not available (requires VScanMagic-GUI)' -Level Info
                }
            } catch {
                $errText = if ($null -eq $_.Exception.Message) { 'Unknown error' } elseif ($_.Exception.Message -is [array]) { ($_.Exception.Message -join '; ') } else { [string]$_.Exception.Message }
                Write-CSApiLog ('Post-download Top X generation failed: ' + $errText) -Level Warning
            }
        }
    }
    }

    return @{ Succeeded = [array]$succeeded; Failed = [array]$failed }
}

# When CS report builder returns 404, skip attempts for remaining reports (reduces failed requests)
$script:CSReportBuilderUnavailable = $false

# Standard report IDs from report_builder/standard_reports - used for asset-level data (Host Name, IP, CVE per row)
$script:CSReportIdMap = @{
    'all-vulnerabilities' = '00000000000000000000000000000000'
    'suppressed-vulnerabilities' = '1d091564830b44c485a0ddc35ace9ac6'
    'external-vulnerabilities' = '01beb6b930744e11b690bb9dc25118fb'
    'executive-summary' = '1cd4f45884264d15bee4173dc58b6a57'
    'pending-epss' = '85d4913c0dbc4fc782b858f0d27dd180'
}
$script:CSReportNameMap = @{
    'all-vulnerabilities' = 'All Vulnerabilities Report'
    'suppressed-vulnerabilities' = 'Suppressed Vulnerabilities'
    'external-vulnerabilities' = 'External Scan'
    'executive-summary' = 'Executive Summary Report'
    'pending-epss' = 'Pending Remediation EPSS Score Reports'
    'network-vulnerabilities' = 'Network Scan Findings'
}

# Report type mapping: our internal type -> possible ConnectSecure report_type/report_id values to try
# ConnectSecure standard reports: All Vulnerabilities, Suppressed Vulnerabilities, External Scan, Executive Summary Report, Pending Remediation EPSS Score Reports
$script:CSReportTypeMap = @{
    'pending-epss' = @('pending_remediation_epss_score_reports','pending_remediation_epss','pending_epss','Pending Remediation EPSS Score Reports')
    'executive-summary' = @('executive_summary_report','executive_summary','Executive Summary Report')
    'all-vulnerabilities' = @('all_vulnerabilities_report','all_vulnerabilities','All Vulnerabilities Report')
    'external-vulnerabilities' = @('external_scan','external_vulnerabilities','External Scan')
    'suppressed-vulnerabilities' = @('suppressed_vulnerabilities','Suppressed Vulnerabilities')
    'network-vulnerabilities' = @('network_scan_findings','network_vulnerabilities','Network Scan Findings')
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
    $t5 = 'registry-vulnerabilities'
    $t6 = 'network-vulnerabilities'
    $t7 = 'registry-remediated'
    $t8 = 'registry-suppressed'
    if ($rt -eq $t0) { New-PendingEPSSReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -VulnerabilityData $vulnDataToPass -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t1) { New-ExecutiveSummaryReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -ScanDate $ScanDate -VulnerabilityData $vulnDataToPass -TopCount $TopCount -DebugLimit $DebugLimit }
    elseif ($rt -eq $t2) { New-AllVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -VulnerabilityData $vulnDataToPass -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t3) { New-ExternalVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t4) { New-SuppressedVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t5) { New-RegistryVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t6) { New-NetworkVulnerabilitiesReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t7) { New-RegistryRemediatedReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    elseif ($rt -eq $t8) { New-RegistrySuppressedReportFromConnectSecure -OutputPath $OutputPath -CompanyId $CompanyId -ClientName $ClientName -OnProgress $OnProgress -DebugLimit $DebugLimit }
    else { throw ('Unknown report type: ' + $InternalReportType) }
}
