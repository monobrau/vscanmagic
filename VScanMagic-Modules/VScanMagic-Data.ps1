# VScanMagic-Data.ps1 - Vulnerability data reading and processing
# Dot-sourced by VScanMagic-GUI.ps1

function Find-ColumnIndex {
    param(
        [hashtable]$Headers,
        [string[]]$PossibleNames
    )

    # Try exact match first (case-insensitive)
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header -eq $name) {
                return $Headers[$header]
            }
        }
    }

    # Try case-insensitive match
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header.ToLower() -eq $name.ToLower()) {
                return $Headers[$header]
            }
        }
    }

    # Try partial match
    foreach ($name in $PossibleNames) {
        foreach ($header in $Headers.Keys) {
            if ($header -like "*$name*" -or $name -like "*$header*") {
                return $Headers[$header]
            }
        }
    }

    return $null
}

function Get-SafeNumericValue {
    param(
        [string]$Value,
        [int]$DefaultValue = 0
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultValue
    }

    # Remove commas and whitespace
    $cleanValue = $Value -replace '[,\s]', ''

    # Try to parse as integer
    $result = 0
    if ([int]::TryParse($cleanValue, [ref]$result)) {
        return $result
    }

    # Try to parse as double and round
    $doubleResult = 0.0
    if ([double]::TryParse($cleanValue, [ref]$doubleResult)) {
        return [int][Math]::Round($doubleResult)
    }

    return $DefaultValue
}

function Get-SafeDoubleValue {
    param(
        [string]$Value,
        [double]$DefaultValue = 0.0
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultValue
    }

    # Remove commas and whitespace
    $cleanValue = $Value -replace '[,\s]', ''

    # Try to parse as double
    $result = 0.0
    if ([double]::TryParse($cleanValue, [ref]$result)) {
        return $result
    }

    return $DefaultValue
}

function Test-SheetMatch {
    param(
        [string]$SheetName,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ($SheetName -like $pattern) {
            return $true
        }
    }
    return $false
}

function Read-SheetData {
    param(
        [object]$Worksheet,
        [hashtable]$ColumnIndices
    )

    $usedRange = $Worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count

    if ($rowCount -le 1) {
        return @()
    }

    Write-Log "  Reading $rowCount rows into memory (bulk read)..."

    # PERFORMANCE OPTIMIZATION: Read entire range into memory with single COM call
    # This is 10-100x faster than reading cells individually
    $rangeValues = $usedRange.Value2

    if ($null -eq $rangeValues) {
        return @()
    }

    Write-Log "  Processing data in memory..."

    # Use ArrayList for better performance than array append
    $data = [System.Collections.ArrayList]::new()

    # Process rows in memory (no COM calls)
    for ($row = 2; $row -le $rowCount; $row++) {
        # Show progress for large datasets
        if ($row % 500 -eq 0) {
            Write-Log "  Processed $row of $rowCount rows..."
        }

        # Get values from 2D array (row, column) - fast, no COM calls
        $product = ''
        if ($columnIndices.ContainsKey('Product')) {
            $product = [string]$rangeValues[$row, $columnIndices['Product']]
        }

        # Skip rows with no product name
        if ([string]::IsNullOrWhiteSpace($product)) {
            continue
        }

        # Build row data from in-memory array
        $hostName = ''
        if ($columnIndices.ContainsKey('HostName')) {
            $hostName = [string]$rangeValues[$row, $columnIndices['HostName']]
        }

        $ip = ''
        if ($columnIndices.ContainsKey('IP')) {
            $ip = [string]$rangeValues[$row, $columnIndices['IP']]
        }

        $username = ''
        if ($columnIndices.ContainsKey('Username')) {
            $username = [string]$rangeValues[$row, $columnIndices['Username']]
        }

        $critical = 0
        if ($columnIndices.ContainsKey('Critical')) {
            $critical = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Critical']])
        }

        $high = 0
        if ($columnIndices.ContainsKey('High')) {
            $high = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['High']])
        }

        $medium = 0
        if ($columnIndices.ContainsKey('Medium')) {
            $medium = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Medium']])
        }

        $low = 0
        if ($columnIndices.ContainsKey('Low')) {
            $low = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['Low']])
        }

        $vulnCount = 0
        if ($columnIndices.ContainsKey('VulnCount')) {
            $vulnCount = Get-SafeNumericValue -Value ([string]$rangeValues[$row, $columnIndices['VulnCount']])
        } else {
            # Calculate from severity counts if not provided
            $vulnCount = $critical + $high + $medium + $low
        }

        $epssScore = 0.0
        if ($columnIndices.ContainsKey('EPSS')) {
            $epssScore = Get-SafeDoubleValue -Value ([string]$rangeValues[$row, $columnIndices['EPSS']])
        }

        # Only add rows that have at least one vulnerability
        if ($vulnCount -gt 0) {
            $null = $data.Add([PSCustomObject]@{
                'Host Name' = $hostName
                'IP' = $ip
                'Username' = $username
                'Product' = $product
                'Critical' = $critical
                'High' = $high
                'Medium' = $medium
                'Low' = $low
                'Vulnerability Count' = $vulnCount
                'EPSS Score' = $epssScore
            })
        }
    }

    Write-Log "  Completed processing $($data.Count) vulnerability records"
    return $data.ToArray()
}

function Test-IsAllVulnerabilitiesFormat {
    param([object]$Workbook)
    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Name -like "*All Vulnerabilities*") {
            Write-Log "Detected All Vulnerabilities single-sheet format" -Level Info
            return $true
        }
    }
    return $false
}

function Test-IsFullListFormat {
    param(
        [object]$Workbook
    )
    
    # Check for sheets that indicate full list format
    $fullListSheetPatterns = @(
        "*Critical Vulnerabilities*",
        "*High Vulnerabilities*",
        "*Medium Vulnerabilities*",
        "*Low Vulnerabilities*",
        "*END OF LIFE*",
        "*End of Life*"
    )
    
    $foundSheets = @()
    foreach ($sheet in $Workbook.Worksheets) {
        foreach ($pattern in $fullListSheetPatterns) {
            if ($sheet.Name -like $pattern) {
                $foundSheets += $sheet.Name
                break
            }
        }
    }
    
    if ($foundSheets.Count -ge 2) {
        Write-Log "Detected full vulnerability list format (found sheets: $($foundSheets -join ', '))" -Level Info
        return $true
    }
    
    return $false
}

function Read-FullListSheetData {
    param(
        [object]$Worksheet,
        [hashtable]$ColumnIndices,
        [switch]$IsEOLSheet,
        [string]$SourceOverride = $null  # When set, use this as Source for all rows (e.g. 'Registry' for Registry worksheet)
    )
    
    $usedRange = $Worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    
    if ($rowCount -le 1) {
        return @()
    }
    
    Write-Log "  Reading $rowCount rows into memory (bulk read)..."
    
    # Read entire range into memory
    $rangeValues = $usedRange.Value2
    
    if ($null -eq $rangeValues) {
        return @()
    }
    
    Write-Log "  Processing data in memory..."
    
    $vulnerabilities = [System.Collections.ArrayList]::new()
    
    # Process rows in memory
    for ($row = 2; $row -le $rowCount; $row++) {
        if ($row % 500 -eq 0) {
            Write-Log "  Processed $row of $rowCount rows..."
        }
        
        # Extract values
        $hostName = ''
        if ($columnIndices.ContainsKey('HostName')) {
            $hostName = [string]$rangeValues[$row, $columnIndices['HostName']]
        }
        
        $ip = ''
        if ($columnIndices.ContainsKey('IP')) {
            $ip = [string]$rangeValues[$row, $columnIndices['IP']]
        }
        
        $product = ''
        if ($columnIndices.ContainsKey('Product')) {
            $product = [string]$rangeValues[$row, $columnIndices['Product']]
            # Clean up product name (remove brackets and quotes if present)
            $product = $product -replace "^\[|'|\]$", "" -replace "^'|'$", ""
            # Tag EOL sheet products so they get max risk weight
            if ($IsEOLSheet -and -not [string]::IsNullOrWhiteSpace($product)) {
                $product = "$product (End of Life)"
            }
        }
        
        $severity = ''
        if ($columnIndices.ContainsKey('Severity')) {
            $severity = [string]$rangeValues[$row, $columnIndices['Severity']]
        }
        
        $epssScore = 0.0
        if ($columnIndices.ContainsKey('EPSS')) {
            $epssScore = Get-SafeDoubleValue -Value ([string]$rangeValues[$row, $columnIndices['EPSS']])
        }
        
        $fix = ''
        if ($columnIndices.ContainsKey('Fix')) {
            $fix = [string]$rangeValues[$row, $columnIndices['Fix']]
            if ([string]::IsNullOrWhiteSpace($fix)) { $fix = '' }
        }
        
        $username = ''
        if ($columnIndices.ContainsKey('Username')) {
            $username = [string]$rangeValues[$row, $columnIndices['Username']]
            if ([string]::IsNullOrWhiteSpace($username)) { $username = '' }
        }
        
        $cve = ''
        if ($columnIndices.ContainsKey('CVE')) {
            $cve = [string]$rangeValues[$row, $columnIndices['CVE']]
            if ([string]::IsNullOrWhiteSpace($cve)) { $cve = '' }
        }
        
        $source = if ($SourceOverride) { $SourceOverride } else { 'Application' }
        if (-not $SourceOverride -and $columnIndices.ContainsKey('Source')) {
            $srcVal = [string]$rangeValues[$row, $columnIndices['Source']]
            if (-not [string]::IsNullOrWhiteSpace($srcVal)) { $source = $srcVal.Trim() }
        }
        
        # Skip rows without required data
        if ([string]::IsNullOrWhiteSpace($product) -or [string]::IsNullOrWhiteSpace($severity)) {
            continue
        }
        
        # Add vulnerability record
        $null = $vulnerabilities.Add([PSCustomObject]@{
            'Source' = $source
            'Host Name' = $hostName
            'IP' = $ip
            'Product' = $product
            'Severity' = $severity
            'EPSS Score' = $epssScore
            'Fix' = $fix
            'Username' = $username
            'CVE' = $cve
        })
    }
    
    Write-Log "  Completed processing $($vulnerabilities.Count) vulnerability records"
    return $vulnerabilities.ToArray()
}

function Get-SeverityCounts {
    param([array]$Items, [string]$SeverityProp = 'Severity', [string]$EPSSProp = 'EPSS Score')
    $critical = 0; $high = 0; $medium = 0; $low = 0; $maxEPSS = 0.0
    foreach ($item in $Items) {
        switch ($item.$SeverityProp) {
            'Critical' { $critical++ }
            'High' { $high++ }
            'Medium' { $medium++ }
            'Low' { $low++ }
        }
        $epss = $item.$EPSSProp
        if ($null -ne $epss -and $epss -gt $maxEPSS) { $maxEPSS = [double]$epss }
    }
    return @{ Critical = $critical; High = $high; Medium = $medium; Low = $low; MaxEPSS = $maxEPSS }
}

function Aggregate-FullListData {
    param(
        [array]$Vulnerabilities
    )
    
    Write-Log "Aggregating $($Vulnerabilities.Count) vulnerabilities by Source/Host/Product/Severity..."
    
    $aggregated = [System.Collections.ArrayList]::new()
    
    # Group by Source, Host Name, IP, and Product (Source keeps Application/Registry/Network in separate sections)
    $grouped = $Vulnerabilities | Group-Object -Property {
        $src = if ($_.Source) { $_.Source } else { 'Application' }
        "$src|$($_.'Host Name')|$($_.IP)|$($_.Product)"
    }
    
    foreach ($group in $grouped) {
        $firstItem = $group.Group[0]
        $sourceVal = if ($firstItem.Source) { $firstItem.Source } else { 'Application' }
        $counts = Get-SeverityCounts -Items $group.Group -EPSSProp 'EPSS Score'
        $vulnCount = $counts.Critical + $counts.High + $counts.Medium + $counts.Low
        # Collect Fix: first non-empty, or concatenate unique if multiple differ
        $fixes = $group.Group | ForEach-Object { $_.Fix } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        $fixVal = if ($fixes.Count -gt 0) { ($fixes -join '; ').Trim() } else { '' }
        # Collect Username: first non-empty from group
        $usernames = $group.Group | ForEach-Object { $_.Username } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        $usernameVal = if ($usernames.Count -gt 0) { [string]$usernames[0] } else { '' }
        # Collect CVE: unique non-empty from group (CVE IDs help AI provide specific remediation)
        $cves = $group.Group | ForEach-Object { $_.CVE } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        $cveVal = if ($cves -and $cves.Count -gt 0) { ($cves -join '; ').Trim() } else { '' }
        
        if ($vulnCount -gt 0) {
            $null = $aggregated.Add([PSCustomObject]@{
                'Source' = $sourceVal
                'Host Name' = $firstItem.'Host Name'
                'IP' = $firstItem.IP
                'Username' = $usernameVal
                'Product' = $firstItem.Product
                'CVE' = $cveVal
                'Critical' = $counts.Critical
                'High' = $counts.High
                'Medium' = $counts.Medium
                'Low' = $counts.Low
                'Vulnerability Count' = $vulnCount
                'EPSS Score' = $counts.MaxEPSS
                'Fix' = $fixVal
            })
        }
    }
    
    Write-Log "Aggregated to $($aggregated.Count) unique Host/Product combinations" -Level Success
    return $aggregated.ToArray()
}

function Get-VulnerabilityData {
    param(
        [string]$ExcelPath
    )

    Write-Log "Reading vulnerability data from Excel..."
    Write-Log "Auto-detecting and consolidating remediation sheets..."

    $excel = $null
    $workbook = $null
    $allData = @()
    $tempPath = $null

    try {
        if (-not (Test-Path -LiteralPath $ExcelPath)) {
            throw "File not found: $ExcelPath"
        }

        # Workaround for OneDrive/sync folder locks - copy to temp before opening
        # Excel COM can fail with "Unable to get the Open property" on cloud-synced paths
        $pathToOpen = $ExcelPath
        if ($ExcelPath -match 'OneDrive|iCloud|Dropbox|Google Drive|Box\.com') {
            $tempDir = [System.IO.Path]::GetTempPath()
            $baseName = [System.IO.Path]::GetFileName($ExcelPath)
            if ([string]::IsNullOrEmpty($baseName)) { $baseName = "vuln_report.xlsx" }
            $tempPath = Join-Path $tempDir ("VScanMagic_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
            Copy-Item -LiteralPath $ExcelPath -Destination $tempPath -Force
            $pathToOpen = $tempPath
            Write-Log "Copied to temp (OneDrive workaround): $tempPath" -Level Info
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # Check for file lock before opening (proactive temp copy if locked)
        if (Test-FileLocked $pathToOpen) {
            if (-not $tempPath) {
                $tempDir = [System.IO.Path]::GetTempPath()
                $baseName = [System.IO.Path]::GetFileName($ExcelPath)
                if ([string]::IsNullOrEmpty($baseName)) { $baseName = "vuln_report.xlsx" }
                $tempPath = Join-Path $tempDir ("VScanMagic_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
                Copy-Item -LiteralPath $ExcelPath -Destination $tempPath -Force
                $pathToOpen = $tempPath
                Write-Log "File is in use - copied to temp: $tempPath" -Level Info
            } else {
                throw "The file is in use by another process. Please close it and try again."
            }
        }

        # UpdateLinks:=0, ReadOnly:=true - helps avoid lock issues
        try {
            $workbook = $excel.Workbooks.Open($pathToOpen, 0, $true)
        } catch {
            # Fallback: if open fails (e.g. sync lock), copy to temp and retry
            if (-not $tempPath) {
                $tempDir = [System.IO.Path]::GetTempPath()
                $baseName = [System.IO.Path]::GetFileName($ExcelPath)
                if ([string]::IsNullOrEmpty($baseName)) { $baseName = "vuln_report.xlsx" }
                $tempPath = Join-Path $tempDir ("VScanMagic_" + [Guid]::NewGuid().ToString("N") + "_" + $baseName)
                Copy-Item -LiteralPath $ExcelPath -Destination $tempPath -Force
                Write-Log "Open failed, retrying from temp copy: $($_.Exception.Message)" -Level Warning
                $workbook = $excel.Workbooks.Open($tempPath, 0, $true)
            } else {
                throw
            }
        }

        # Log all sheets found in workbook for debugging
        Write-Log "All sheets found in workbook:"
        $allSheetNames = @()
        foreach ($sheet in $workbook.Worksheets) {
            $allSheetNames += $sheet.Name
            Write-Log "  - '$($sheet.Name)'"
        }
        Write-Log "Total sheets: $($allSheetNames.Count)"
        
        # Check if this is an All Vulnerabilities single-sheet format (ConnectSecure 13-column)
        $isAllVulnsFormat = Test-IsAllVulnerabilitiesFormat -Workbook $workbook
        if ($isAllVulnsFormat) {
            # All Vulnerabilities report can have multiple worksheets: main + Registry + Network Scan Findings
            $avSheetPatterns = @('*All Vulnerabilities*', '*Registry*', '*Network*')
            $sheetsToRead = @()
            foreach ($sheet in $workbook.Worksheets) {
                foreach ($pat in $avSheetPatterns) {
                    if ($sheet.Name -like $pat) {
                        $src = $null
                        if ($pat -like '*Registry*') { $src = 'Registry' }
                        elseif ($pat -like '*Network*') { $src = 'Network' }
                        $sheetsToRead += @{ Sheet = $sheet; Source = $src }
                        break
                    }
                }
            }
            if ($sheetsToRead.Count -eq 0) {
                throw "All Vulnerabilities sheet not found."
            }
            # Use first sheet for headers (All Vulnerabilities or Registry - both use same 13-column structure)
            $firstEntry = $sheetsToRead[0]
            $firstSheet = $firstEntry.Sheet
            $usedRange = $firstSheet.UsedRange
            $colCount = $usedRange.Columns.Count
            $headers = @{}
            for ($col = 1; $col -le $colCount; $col++) {
                $headerName = $firstSheet.Cells.Item(1, $col).Text
                if ($headerName) {
                    $headers[$headerName] = $col
                }
            }
            Write-Log "All Vulnerabilities headers: $($headers.Keys -join ', ')"
            $columnMappings = @{
                'Source' = @('Source', 'Section', 'Vulnerability Source')
                'HostName' = @('Asset Name', 'Host Name', 'Hostname', 'Computer', 'Device')
                'IP' = @('IP Address', 'IP', 'Address')
                'Product' = @('Product Name', 'Application Name', 'Software Name', 'Product', 'App Name', 'OS Name', 'OS Full Name', 'Name', 'Problem Name')
                'Severity' = @('Severity')
                'EPSS' = @('EPSS Score', 'EPSS', 'Exploit Prediction Score')
                'Fix' = @('Solution', 'Fix', 'Remediation', 'FIX')
                'Username' = @('Username', 'User Name', 'User', 'Account', 'Login', 'Login Name', 'Last User', 'Last Logged In User', 'Last Logged In User Name', 'Logged In User', 'Owner', 'Asset Owner', 'Primary User')
                'CVE' = @('CVE ID', 'CVE', 'Problem Name', 'problem_name')
            }
            $columnIndices = @{}
            foreach ($key in $columnMappings.Keys) {
                $colIndex = Find-ColumnIndex -Headers $headers -PossibleNames $columnMappings[$key]
                if ($colIndex) {
                    $columnIndices[$key] = $colIndex
                }
            }
            if (-not $columnIndices.ContainsKey('Product') -or -not $columnIndices.ContainsKey('Severity')) {
                throw "All Vulnerabilities format requires Product Name (or Application Name) and Severity columns."
            }
            $sheetVulns = @()
            foreach ($entry in $sheetsToRead) {
                $s = $entry.Sheet
                $srcOverride = $entry.Source
                $vulns = Read-FullListSheetData -Worksheet $s -ColumnIndices $columnIndices -IsEOLSheet:$false -SourceOverride $srcOverride
                Write-Log "Read $($vulns.Count) vulnerabilities from $($s.Name)"
                $sheetVulns += $vulns
                Clear-ComObject $s
            }
            Clear-ComObject $usedRange
            $sheetVulns = Expand-CombinedProductRows -Vulnerabilities $sheetVulns
            $allData = Aggregate-FullListData -Vulnerabilities $sheetVulns
            Write-Log "Aggregated to $($allData.Count) records" -Level Success
            return $allData
        }

        # Check if this is a full list format file (multi-sheet Critical/High/Medium/Low)
        $isFullListFormat = Test-IsFullListFormat -Workbook $workbook
        
        if ($isFullListFormat) {
            Write-Log "Processing as full vulnerability list format..." -Level Info
            
            # Find vulnerability severity sheets (includes END OF LIFE - treated as max risk)
            $fullListSheetPatterns = @(
                "*Critical Vulnerabilities*",
                "*High Vulnerabilities*",
                "*Medium Vulnerabilities*",
                "*Low Vulnerabilities*",
                "*END OF LIFE*",
                "*End of Life*"
            )
            
            $sourceSheets = @()
            foreach ($sheet in $workbook.Worksheets) {
                $sheetName = $sheet.Name
                foreach ($pattern in $fullListSheetPatterns) {
                    if ($sheetName -like $pattern) {
                        Write-Log "Found vulnerability list sheet: $sheetName"
                        $sourceSheets += $sheet
                        break
                    }
                }
            }
            
            if ($sourceSheets.Count -eq 0) {
                throw "No vulnerability list sheets found. Expected sheets matching: Critical Vulnerabilities, High Vulnerabilities, Medium Vulnerabilities, Low Vulnerabilities"
            }
            
            # Get headers from first sheet
            $firstSheet = $sourceSheets[0]
            $usedRange = $firstSheet.UsedRange
            $colCount = $usedRange.Columns.Count
            
            # Get headers
            $headers = @{}
            for ($col = 1; $col -le $colCount; $col++) {
                $headerName = $firstSheet.Cells.Item(1, $col).Text
                if ($headerName) {
                    $headers[$headerName] = $col
                }
            }
            
            Write-Log "Found headers: $($headers.Keys -join ', ')"
            
            # Define column mappings for full list format
            $columnMappings = @{
                'HostName' = @('Host Name', 'Hostname', 'Computer', 'Computer Name', 'Device', 'Device Name', 'System', 'System Name', 'Machine')
                'IP' = @('IP', 'IP Address', 'IPAddress', 'Address')
                'Product' = @('Software Name', 'Product', 'Software', 'Application', 'App', 'Program', 'Title', 'Product Name', 'OS Name', 'OS Full Name', 'Name', 'Problem Name')
                'Severity' = @('Severity')
                'EPSS' = @('EPSS Score', 'EPSS', 'Exploit Prediction Score')
                'Fix' = @('Solution', 'Fix', 'Remediation', 'FIX')
                'Username' = @('Username', 'User Name', 'User', 'Account', 'Login', 'Login Name', 'Last User', 'Last Logged In User', 'Last Logged In User Name', 'Logged In User', 'Owner', 'Asset Owner', 'Primary User')
                'CVE' = @('CVE ID', 'CVE', 'Problem Name', 'problem_name')
            }
            
            # Find column indices
            $columnIndices = @{}
            foreach ($key in $columnMappings.Keys) {
                $colIndex = Find-ColumnIndex -Headers $headers -PossibleNames $columnMappings[$key]
                if ($colIndex) {
                    $columnIndices[$key] = $colIndex
                    Write-Log "Mapped '$key' to column: $($headers.Keys | Where-Object { $headers[$_] -eq $colIndex })"
                } else {
                    Write-Log "Could not find column for '$key' (tried: $($columnMappings[$key] -join ', '))" -Level Warning
                }
            }
            
            # Verify required columns
            $requiredFields = @('Product', 'Severity')
            $missingRequired = @()
            foreach ($field in $requiredFields) {
                if (-not $columnIndices.ContainsKey($field)) {
                    $missingRequired += $field
                }
            }
            
            if ($missingRequired.Count -gt 0) {
                throw "Missing required columns for full list format: $($missingRequired -join ', '). Please ensure your Excel file has Product/Software Name and Severity columns."
            }
            
            Write-Log "Successfully mapped $($columnIndices.Count) columns."
            
            # Read all vulnerabilities from all sheets
            $allVulnerabilities = @()
            foreach ($sheet in $sourceSheets) {
                Write-Log "Reading vulnerabilities from: $($sheet.Name)"
                $isEOLSheet = $sheet.Name -like "*END OF LIFE*" -or $sheet.Name -like "*End of Life*"
                $sheetVulns = Read-FullListSheetData -Worksheet $sheet -ColumnIndices $columnIndices -IsEOLSheet:$isEOLSheet
                Write-Log "  Found $($sheetVulns.Count) vulnerabilities"
                $allVulnerabilities += $sheetVulns
            }
            
            Write-Log "Total vulnerabilities read: $($allVulnerabilities.Count)"
            
            # Expand combined products (CS often combines e.g. MongoDB 3.4.x or Visual C++ 2008/2010/2012 in one cell)
            $allVulnerabilities = Expand-CombinedProductRows -Vulnerabilities $allVulnerabilities
            # Aggregate vulnerabilities by Host/Product/Severity
            $allData = Aggregate-FullListData -Vulnerabilities $allVulnerabilities
            
            Write-Log "Total vulnerability records consolidated: $($allData.Count)" -Level Success
            
            # Clean up sheet references
            foreach ($sheet in $sourceSheets) {
                Clear-ComObject $sheet
            }
            
            return $allData
        }
        
        # Original aggregated format processing
        Write-Log "Processing as aggregated format..." -Level Info
        
        # Find all sheets that match remediation patterns
        $sourceSheets = @()
        
        foreach ($sheet in $workbook.Worksheets) {
            $sheetName = $sheet.Name

            # Skip excluded sheets
            $shouldExclude = $false
            foreach ($excludePattern in $script:Config.ExcludeSheetPatterns) {
                if ($sheetName -like $excludePattern -or $sheetName -eq $excludePattern) {
                    $shouldExclude = $true
                    break
                }
            }

            if ($shouldExclude) {
                Write-Log "Excluding sheet: $sheetName"
                Clear-ComObject $sheet
                continue
            }

            # Check if sheet matches any remediation pattern
            $isMatch = Test-SheetMatch -SheetName $sheetName -Patterns $script:Config.SourceSheetPatterns

            if ($isMatch) {
                Write-Log "Found remediation sheet: $sheetName"
                $sourceSheets += $sheet
            } else {
                Write-Log "Sheet '$sheetName' does not match remediation patterns (looking for: $($script:Config.SourceSheetPatterns -join ', '))"
                Clear-ComObject $sheet
            }
        }

        if ($sourceSheets.Count -eq 0) {
            throw "No remediation sheets found. Looking for patterns: $($script:Config.SourceSheetPatterns -join ', '). Excluding: $($script:Config.ExcludeSheetPatterns -join ', ')"
        }

        Write-Log "Processing $($sourceSheets.Count) remediation sheet(s)..."

        # Get headers from first sheet and create column mappings
        $firstSheet = $sourceSheets[0]
        $usedRange = $firstSheet.UsedRange
        $colCount = $usedRange.Columns.Count

        # Get headers
        $headers = @{}
        for ($col = 1; $col -le $colCount; $col++) {
            $headerName = $firstSheet.Cells.Item(1, $col).Text
            if ($headerName) {
                $headers[$headerName] = $col
            }
        }

        Write-Log "Found headers: $($headers.Keys -join ', ')"

        # Define flexible column mappings
        $columnMappings = @{
            'HostName' = @('Host Name', 'Hostname', 'Computer', 'Computer Name', 'Device', 'Device Name', 'System', 'System Name', 'Machine')
            'IP' = @('IP', 'IP Address', 'IPAddress', 'Address')
            'Username' = @('Username', 'User Name', 'User', 'Account', 'Login', 'Login Name', 'Last User', 'Last Logged In User', 'Last Logged In User Name', 'Logged In User', 'Owner', 'Asset Owner', 'Primary User')
            'Product' = @('Product', 'Software', 'Application', 'App', 'Program', 'Title', 'Product Name', 'Software Name')
            'Critical' = @('Critical', 'Crit', 'Critical Count', 'Critical Vulnerabilities')
            'High' = @('High', 'High Count', 'High Vulnerabilities')
            'Medium' = @('Medium', 'Med', 'Medium Count', 'Medium Vulnerabilities')
            'Low' = @('Low', 'Low Count', 'Low Vulnerabilities')
            'VulnCount' = @('Vulnerability Count', 'Vuln Count', 'Total Vulnerabilities', 'Total Vulns', 'Count', 'Total Count', 'Number of Vulnerabilities')
            'EPSS' = @('EPSS Score', 'EPSS', 'Exploit Prediction Score', 'Max EPSS Score', 'Max EPSS')
        }

        # Find column indices
        $columnIndices = @{}
        foreach ($key in $columnMappings.Keys) {
            $colIndex = Find-ColumnIndex -Headers $headers -PossibleNames $columnMappings[$key]
            if ($colIndex) {
                $columnIndices[$key] = $colIndex
                Write-Log "Mapped '$key' to column: $($headers.Keys | Where-Object { $headers[$_] -eq $colIndex })"
            } else {
                Write-Log "Could not find column for '$key' (tried: $($columnMappings[$key] -join ', '))" -Level Warning
            }
        }

        # Verify minimum required columns
        $requiredFields = @('Product')
        $missingRequired = @()
        foreach ($field in $requiredFields) {
            if (-not $columnIndices.ContainsKey($field)) {
                $missingRequired += $field
            }
        }

        if ($missingRequired.Count -gt 0) {
            throw "Missing required columns: $($missingRequired -join ', '). Please ensure your Excel file has at least a Product/Software column."
        }

        Write-Log "Successfully mapped $($columnIndices.Count) columns."

        # Read data from all matching sheets
        foreach ($sheet in $sourceSheets) {
            Write-Log "Reading data from: $($sheet.Name)"
            $sheetData = Read-SheetData -Worksheet $sheet -ColumnIndices $columnIndices
            Write-Log "  Found $($sheetData.Count) vulnerability records"
            $allData += $sheetData
        }

        Write-Log "Total vulnerability records consolidated: $($allData.Count)" -Level Success

        # Clean up sheet references
        foreach ($sheet in $sourceSheets) {
            Clear-ComObject $sheet
        }

        return $allData

    } catch {
        Write-Log "Error reading Excel data: $($_.Exception.Message)" -Level Error
        throw
    } finally {
        if ($workbook) {
            $workbook.Close($false)
            Clear-ComObject $workbook
        }
        if ($excel) {
            $excel.Quit()
            Clear-ComObject $excel
        }
        if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-ProductMajorVersion {
    <#
    .SYNONOPSIS
    Normalizes product name to major version for report formatting. ConnectSecure often combines
    many minor versions (e.g. MongoDB 3.4.24, 3.4.23, 3.4.22) - this consolidates to "MongoDB 3.4".
    #>
    param([string]$ProductName)
    if ([string]::IsNullOrWhiteSpace($ProductName)) { return $ProductName }
    $p = $ProductName.Trim()
    # Strip trailing patch version: "MongoDB 3.4.24" -> "MongoDB 3.4", "Product 9.0.30729.4974" -> "Product"
    if ($p -match '^(.+?)\s+(\d+\.\d+)(\.\d+)+$') {
        $base = $matches[1].Trim()
        $majorMinor = $matches[2]
        return "$base $majorMinor"
    }
    # Strip trailing single version: "Product 10.0.40219" -> "Product"
    if ($p -match '^(.+?)\s+\d+\.\d+(\.\d+)*$') {
        return $matches[1].Trim()
    }
    return $p
}

function Expand-CombinedProductRows {
    <#
    .SYNOPSIS
    Splits rows where Product contains comma/semicolon-separated values (ConnectSecure combines these).
    Each part is normalized to major version. Returns expanded array of rows.
    #>
    param([array]$Vulnerabilities)
    if (-not $Vulnerabilities -or $Vulnerabilities.Count -eq 0) { return $Vulnerabilities }
    $expanded = [System.Collections.ArrayList]::new()
    foreach ($v in $Vulnerabilities) {
        $product = if ($v.Product) { [string]$v.Product } else { '' }
        if ([string]::IsNullOrWhiteSpace($product) -or ($product -notmatch '[,;]')) {
            $null = $expanded.Add($v)
            continue
        }
        $parts = $product -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if ($parts.Count -le 1) {
            $null = $expanded.Add($v)
            continue
        }
        # If first part has product name, prepend to version-only parts (e.g. "MongoDB 3.4.24, 3.4.23" -> "MongoDB 3.4.23")
        $productBase = $null
        if ($parts[0] -match '^(.+?)\s+\d') { $productBase = $matches[1].Trim() }
        foreach ($part in $parts) {
            $resolved = $part
            if ($productBase -and $part -match '^\d+\.\d') { $resolved = "$productBase $part" }
            $normProduct = Get-ProductMajorVersion -ProductName $resolved
            $ht = @{}
            $v.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
            $ht['Product'] = $normProduct
            $null = $expanded.Add([PSCustomObject]$ht)
        }
    }
    return $expanded.ToArray()
}

function Get-ConsolidatedProduct {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $ProductName
    }

    # Normalize the product name for comparison
    $normalizedProduct = $ProductName.Trim()

    # Consolidate Visual C++ and .NET variants early (so Top N includes more diverse items)
    $groupKey = Get-TimeEstimateGroupKey -ProductName $normalizedProduct
    if ($groupKey -ne $normalizedProduct) {
        return $groupKey
    }

    # Check against consolidation rules (case-insensitive)
    foreach ($consolidated in $script:Config.WindowsConsolidation.Keys) {
        $patterns = $script:Config.WindowsConsolidation[$consolidated]
        foreach ($pattern in $patterns) {
            # Try exact match (case-insensitive)
            if ($normalizedProduct -eq $pattern) {
                return $consolidated
            }
            # Try wildcard match
            if ($normalizedProduct -like "*$pattern*") {
                return $consolidated
            }
        }
    }

    # Additional common product normalization
    # Remove version numbers at the end (e.g., "Adobe Reader 11.0.23" -> "Adobe Reader")
    if ($normalizedProduct -match '^(.+?)\s+[\d\.]+$') {
        $baseProduct = $matches[1]
        # Check if this base product should be consolidated
        foreach ($consolidated in $script:Config.WindowsConsolidation.Keys) {
            $patterns = $script:Config.WindowsConsolidation[$consolidated]
            foreach ($pattern in $patterns) {
                if ($baseProduct -like "*$pattern*") {
                    return $consolidated
                }
            }
        }
    }

    return $ProductName
}

function Test-IsMicrosoftApplication {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    # List of Microsoft application patterns (not OS components)
    $microsoftAppPatterns = @(
        'Microsoft Office',
        'Microsoft 365',
        'Microsoft Teams',
        'Microsoft Edge',
        'Microsoft OneDrive',
        'Microsoft Outlook',
        'Microsoft Word',
        'Microsoft Excel',
        'Microsoft PowerPoint',
        'Microsoft Access',
        'Microsoft Publisher',
        'Microsoft Visio',
        'Microsoft Project',
        'Microsoft SharePoint',
        'Skype for Business',
        'Microsoft SQL Server Management Studio',
        'Microsoft Visual Studio Code',
        'Microsoft .NET Framework',
        'Microsoft .NET Core',
        'Microsoft .NET Runtime'
    )

    foreach ($pattern in $microsoftAppPatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Test-IsVMwareProduct {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    # List of VMware product patterns
    $vmwarePatterns = @(
        'VMware',
        'VMWare',
        'vSphere',
        'vCenter',
        'ESXi',
        'VMware Tools',
        'VMware Workstation',
        'VMware Player',
        'VMware Horizon',
        'vRealize',
        'vCloud',
        'NSX'
    )

    foreach ($pattern in $vmwarePatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Test-IsAutoUpdatingSoftware {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    # List of software that auto-updates
    $autoUpdatePatterns = @(
        'Google Chrome',
        'Mozilla Firefox'
    )

    foreach ($pattern in $autoUpdatePatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Get-ProductTypeSuffix {
    <#
    .SYNOPSIS
    Returns the product-type suffix for ticket subjects and report titles (e.g. " - Updates Required").
    Used by New-TicketInstructions, New-TicketInstructionsHtml, and New-WordReport.
    #>
    param(
        [string]$ProductName,
        [bool]$IsRMITPlus = $false
    )
    if ([string]::IsNullOrWhiteSpace($ProductName)) { return " - Update Required" }
    if ($ProductName -like "*Windows Server 2012*" -or $ProductName -like "*end-of-life*" -or $ProductName -like "*out of support*") {
        return " - End of Support Migration Required"
    }
    if ($ProductName -like "*Windows 10*") { return " - Windows 10 is End of Life" }
    if ($ProductName -like "*Windows 11*") { return " - Updates Required" }
    if ($ProductName -like "*Windows Server*") { return " - Updates Required" }
    if ($ProductName -like "*Windows*") { return " - Patch Management Required" }
    if ($ProductName -like "*printer*" -or $ProductName -like "*Ripple20*") { return " - Firmware Update Required" }
    if ($ProductName -like "*Microsoft Teams*") { return " - Application Update Required" }
    if ((Test-IsMicrosoftApplication -ProductName $ProductName) -and $IsRMITPlus) { return " - Updates Required" }
    if ((Test-IsVMwareProduct -ProductName $ProductName) -and $IsRMITPlus) { return " - Update Required" }
    if (Test-IsAutoUpdatingSoftware -ProductName $ProductName) { return " - This software updates automatically" }
    return " - Update Required"
}

function Get-AverageCVSS {
    param(
        [int]$Critical,
        [int]$High,
        [int]$Medium,
        [int]$Low
    )

    $total = $Critical + $High + $Medium + $Low

    if ($total -eq 0) {
        return 0
    }

    $weighted = ($Critical * $script:Config.CVSSEquivalent.Critical) +
                ($High * $script:Config.CVSSEquivalent.High) +
                ($Medium * $script:Config.CVSSEquivalent.Medium) +
                ($Low * $script:Config.CVSSEquivalent.Low)

    return [Math]::Round($weighted / $total, 2)
}

function Test-IsEOLProduct {
    param([string]$ProductName)

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $false
    }

    foreach ($pattern in $script:Config.EOLProductPatterns) {
        if ($ProductName -like "*$pattern*") {
            return $true
        }
    }
    return $false
}

function Get-CompositeRiskScore {
    param(
        [int]$Critical,
        [int]$High,
        [int]$Medium,
        [int]$Low,
        [double]$EPSSScore,
        [string]$ProductName = '',
        [int]$VulnCount = 0
    )

    # ConnectSecure-aligned risk score formula:
    # Severity weighted sum (from Problem Category Weightage) × (1 + EPSS)
    # EOL gets maximum weight (1.0) per ConnectSecure - "maximum-risk scoring event"
    # Severity weights: Critical 0.90, High 0.80, Medium 0.50, Low 0.30
    $severityWeightedSum = ($Critical * $script:Config.SeverityWeights.Critical) +
                           ($High * $script:Config.SeverityWeights.High) +
                           ($Medium * $script:Config.SeverityWeights.Medium) +
                           ($Low * $script:Config.SeverityWeights.Low)

    # EOL products get max weight (1.0) per vuln - hits CS score hard
    if (Test-IsEOLProduct -ProductName $ProductName) {
        $severityWeightedSum += ($VulnCount * 1.0)
    }

    # EPSS boost: (1 + EPSS) ranges from 1.0 to 2.0
    # Items without EPSS (e.g. Network vulns) get synthetic EPSS so they're not buried
    $effectiveEPSS = $EPSSScore
    if ($null -eq $effectiveEPSS -or $effectiveEPSS -eq '' -or ($effectiveEPSS -as [double]) -eq $null) {
        $effectiveEPSS = if ($null -ne $script:Config.SyntheticEPSSForNoEPSS) { $script:Config.SyntheticEPSSForNoEPSS } else { 0.1 }
    }
    $epssFactor = 1.0 + $effectiveEPSS
    $riskScore = $severityWeightedSum * $epssFactor

    return [Math]::Round($riskScore, 2)
}

function Get-Top10Vulnerabilities {
    param(
        [array]$VulnData,
        [double]$MinEPSS = 0,
        [bool]$IncludeCritical = $true,
        [bool]$IncludeHigh = $true,
        [bool]$IncludeMedium = $true,
        [bool]$IncludeLow = $true,
        [int]$Count = 10
    )

    $countText = if ($Count -le 0) { "all" } else { "$Count" }
    Write-Log "Calculating risk scores and identifying top $countText vulnerabilities..."
    Write-Log "Filters: MinEPSS=$MinEPSS, Critical=$IncludeCritical, High=$IncludeHigh, Medium=$IncludeMedium, Low=$IncludeLow, Count=$countText"

    # Group by Source and Product (keeps Application/Registry/Network in separate sections)
    $grouped = $VulnData | Group-Object -Property { $src = if ($_.Source) { $_.Source } else { 'Application' }; "$src|$($_.Product)" }

    $aggregated = @()

    foreach ($group in $grouped) {
        $firstItem = $group.Group[0]
        $sourceVal = if ($firstItem.Source) { $firstItem.Source } else { 'Application' }
        $product = $firstItem.Product

        # Check if product should be filtered
        $shouldFilter = $false
        foreach ($filter in $script:Config.FilteredProducts) {
            if ($product -like "*$filter*") {
                $shouldFilter = $true
                break
            }
        }

        if ($shouldFilter) {
            Write-Log "Filtering out: $product"
            continue
        }

        # Consolidate Windows versions
        $consolidatedProduct = Get-ConsolidatedProduct -ProductName $product

        # Check if we already have this consolidated product (within same Source)
        $existing = $aggregated | Where-Object { $_.Source -eq $sourceVal -and $_.Product -eq $consolidatedProduct }

        if ($existing) {
            # Merge with existing
            $existing.Critical += ($group.Group | Measure-Object -Property Critical -Sum).Sum
            $existing.High += ($group.Group | Measure-Object -Property High -Sum).Sum
            $existing.Medium += ($group.Group | Measure-Object -Property Medium -Sum).Sum
            $existing.Low += ($group.Group | Measure-Object -Property Low -Sum).Sum
            $existing.VulnCount += ($group.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum

            # Take max EPSS score
            $maxEPSS = ($group.Group.'EPSS Score' | Measure-Object -Maximum).Maximum
            if ($maxEPSS -gt $existing.EPSSScore) {
                $existing.EPSSScore = $maxEPSS
            }

            # Collect Fix: use first non-empty from group, or append unique if existing has none
            $groupFix = ($group.Group | ForEach-Object { $_.Fix } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join '; '
            if ($groupFix -and [string]::IsNullOrWhiteSpace($existing.Fix)) { $existing.Fix = $groupFix.Trim() }
            elseif ($groupFix -and $existing.Fix -notlike "*$groupFix*") { $existing.Fix = ($existing.Fix + '; ' + $groupFix).Trim() }
            # Collect CVE: append unique from group
            $groupCves = $group.Group | ForEach-Object { $_.CVE } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            if ($groupCves -and $groupCves.Count -gt 0) {
                $newCves = ($groupCves -join '; ').Trim()
                $existingCveIds = if ($existing.CveIds) { $existing.CveIds } else { '' }
                if ([string]::IsNullOrWhiteSpace($existingCveIds)) { $existing.CveIds = $newCves }
                else {
                    $existingSet = $existingCveIds -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $merged = ($existingSet + ($newCves -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) | Select-Object -Unique
                    $existing.CveIds = ($merged -join '; ').Trim()
                }
            }

            # Add affected systems (store objects with hostname, IP, username, and vulnerability count)
            # Group by Host+IP composite so we capture ALL unique systems (hostname or IP fallback)
            $hostKeyGroups = $group.Group | Group-Object -Property { "$($_.'Host Name')`t$($_.IP)" }
            foreach ($hostGroup in $hostKeyGroups) {
                $hostItem = $hostGroup.Group[0]
                $hostVulnCount = ($hostGroup.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum
                $existing.AffectedSystems += [PSCustomObject]@{
                    HostName = $hostItem.'Host Name'
                    IP = $hostItem.'IP'
                    Username = $hostItem.'Username'
                    VulnCount = $hostVulnCount
                }
            }
        } else {
            # Create new entry
            $critical = ($group.Group | Measure-Object -Property Critical -Sum).Sum
            $high = ($group.Group | Measure-Object -Property High -Sum).Sum
            $medium = ($group.Group | Measure-Object -Property Medium -Sum).Sum
            $low = ($group.Group | Measure-Object -Property Low -Sum).Sum
            $vulnCount = ($group.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum
            $epssScore = ($group.Group.'EPSS Score' | Measure-Object -Maximum).Maximum

            # Collect Fix: first non-empty or concatenate unique from group
            $fixes = $group.Group | ForEach-Object { $_.Fix } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            $fixVal = if ($fixes.Count -gt 0) { ($fixes -join '; ').Trim() } else { '' }
            # Collect CVE: unique non-empty from group
            $cves = $group.Group | ForEach-Object { $_.CVE } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            $cveIds = if ($cves -and $cves.Count -gt 0) { ($cves -join '; ').Trim() } else { '' }

            $avgCVSS = Get-AverageCVSS -Critical $critical -High $high -Medium $medium -Low $low
            $riskScore = Get-CompositeRiskScore -Critical $critical -High $high -Medium $medium -Low $low -EPSSScore $epssScore -ProductName $consolidatedProduct -VulnCount $vulnCount

            # Create affected systems array with hostname, IP, username, and vulnerability count
            # Group by Host+IP composite so we capture ALL unique systems (hostname or IP fallback)
            $affectedSystems = @()
            $hostKeyGroups = $group.Group | Group-Object -Property { "$($_.'Host Name')`t$($_.IP)" }
            foreach ($hostGroup in $hostKeyGroups) {
                $hostItem = $hostGroup.Group[0]
                $hostVulnCount = ($hostGroup.Group | Measure-Object -Property 'Vulnerability Count' -Sum).Sum
                $affectedSystems += [PSCustomObject]@{
                    HostName = $hostItem.'Host Name'
                    IP = $hostItem.'IP'
                    Username = $hostItem.'Username'
                    VulnCount = $hostVulnCount
                }
            }

            $aggregated += [PSCustomObject]@{
                Source = $sourceVal
                Product = $consolidatedProduct
                Critical = $critical
                High = $high
                Medium = $medium
                Low = $low
                VulnCount = $vulnCount
                EPSSScore = $epssScore
                AvgCVSS = $avgCVSS
                RiskScore = $riskScore
                AffectedSystems = $affectedSystems
                Fix = $fixVal
                CveIds = $cveIds
            }
        }
    }

    # Recalculate scores for consolidated entries
    foreach ($item in $aggregated) {
        $item.AvgCVSS = Get-AverageCVSS -Critical $item.Critical -High $item.High -Medium $item.Medium -Low $item.Low
        $item.RiskScore = Get-CompositeRiskScore -Critical $item.Critical -High $item.High -Medium $item.Medium -Low $item.Low -EPSSScore $item.EPSSScore -ProductName $item.Product -VulnCount $item.VulnCount
    }

    # Apply filters
    Write-Log "Before filtering: $($aggregated.Count) products"

    # Count products by severity in single pass
    $productsWithCritical = $productsWithHigh = $productsWithMedium = $productsWithLow = 0
    foreach ($p in $aggregated) {
        if ($p.Critical -gt 0) { $productsWithCritical++ }
        if ($p.High -gt 0) { $productsWithHigh++ }
        if ($p.Medium -gt 0) { $productsWithMedium++ }
        if ($p.Low -gt 0) { $productsWithLow++ }
    }
    Write-Log "Products by severity: Critical=$productsWithCritical, High=$productsWithHigh, Medium=$productsWithMedium, Low=$productsWithLow"
    
    $filtered = $aggregated | Where-Object {
        # EPSS filter: only include items with effective EPSS >= MinEPSS
        # Items without valid EPSS (e.g. Network vulns) get synthetic EPSS for this check
        $hasEpss = $null -ne $_.EPSSScore -and $_.EPSSScore -ne '' -and ($_.EPSSScore -as [double]) -ne $null
        $syn = if ($null -ne $script:Config.SyntheticEPSSForNoEPSS) { $script:Config.SyntheticEPSSForNoEPSS } else { 0.1 }
        $effectiveEPSS = if ($hasEpss) { [double]$_.EPSSScore } else { $syn }
        $epssPass = $effectiveEPSS -ge $MinEPSS
        
        # Severity filter - include if has any of the selected severities
        $severityPass = $false
        if ($IncludeCritical -and $_.Critical -gt 0) { $severityPass = $true }
        if ($IncludeHigh -and $_.High -gt 0) { $severityPass = $true }
        if ($IncludeMedium -and $_.Medium -gt 0) { $severityPass = $true }
        if ($IncludeLow -and $_.Low -gt 0) { $severityPass = $true }
        
        return ($epssPass -and $severityPass)
    }

    Write-Log "Filtered from $($aggregated.Count) to $($filtered.Count) products matching criteria"

    $filteredCritical = $filteredHigh = $filteredMedium = $filteredLow = 0
    foreach ($p in $filtered) {
        if ($p.Critical -gt 0) { $filteredCritical++ }
        if ($p.High -gt 0) { $filteredHigh++ }
        if ($p.Medium -gt 0) { $filteredMedium++ }
        if ($p.Low -gt 0) { $filteredLow++ }
    }
    Write-Log "Filtered products by severity: Critical=$filteredCritical, High=$filteredHigh, Medium=$filteredMedium, Low=$filteredLow"

    # Combined pool: any item with both CVSS and EPSS scores (Application, Registry, Network, or other - if they have both).
    # Separate pool: items lacking either score (e.g. Network when EPSS is empty).
    $hasValidEPSS = {
        $v = $_.EPSSScore
        $null -ne $v -and $v -ne '' -and ($v -as [double]) -ne $null
    }
    $hasCVSS = { $_.VulnCount -gt 0 -or ($null -ne $_.AvgCVSS -and $_.AvgCVSS -gt 0) }
    $combinedPool = $filtered | Where-Object { (& $hasValidEPSS) -and (& $hasCVSS) }
    $separatePool = $filtered | Where-Object { -not ((& $hasValidEPSS) -and (& $hasCVSS)) }

    # Combine both pools, sort by RiskScore, then take top N total (not N per pool)
    $allSorted = @($combinedPool) + @($separatePool) | Sort-Object -Property RiskScore -Descending
    $topVulns = if ($Count -le 0) { $allSorted } else { $allSorted | Select-Object -First $Count }
    $countMsg = if ($Count -le 0) { "all" } else { "top $Count (CVSS+EPSS items combined, others separate)" }

    # Ensure at least MinNetworkVulnsInTopN Network vulns are included (regardless of score)
    $minNetwork = if ($null -ne $script:Config.MinNetworkVulnsInTopN) { $script:Config.MinNetworkVulnsInTopN } else { 2 }
    $networkInTop = @($topVulns | Where-Object { $_.Source -eq 'Network' })
    if ($networkInTop.Count -lt $minNetwork) {
        $networkPool = $filtered | Where-Object { $_.Source -eq 'Network' } | Sort-Object -Property RiskScore -Descending
        $topVulnKeys = $topVulns | ForEach-Object { "$($_.Source)|$($_.Product)" }
        $needed = $minNetwork - $networkInTop.Count
        $added = 0
        foreach ($n in $networkPool) {
            if ($added -ge $needed) { break }
            $key = "$($n.Source)|$($n.Product)"
            if ($key -notin $topVulnKeys) {
                $topVulns += $n
                $topVulnKeys += $key
                $added++
            }
        }
        if ($added -gt 0) {
            $topVulns = $topVulns | Sort-Object -Property RiskScore -Descending
            Write-Log "Added $added Network vuln(s) to meet minimum of $minNetwork (total now $($topVulns.Count))" -Level Info
        }
    }

    Write-Log "Identified $($topVulns.Count) vulnerabilities ($countMsg)" -Level Success

    return $topVulns
}

function Get-RemediationGuidance {
    param(
        [string]$ProductName,
        [ValidateSet('Word', 'Ticket')]
        [string]$OutputType
    )

    if ($null -eq $script:RemediationRules -or $script:RemediationRules.Count -eq 0) {
        Load-RemediationRules
    }

    # Use cached sorted rules (invalidated on Load/Save)
    if ($null -eq $script:CachedRemediationRulesForGuidance) {
        $nonDefault = @($script:RemediationRules | Where-Object { -not $_.IsDefault -and $_.Pattern -ne "*" } | Sort-Object { $_.Pattern.Length } -Descending)
        $default = $script:RemediationRules | Where-Object { $_.IsDefault -or $_.Pattern -eq "*" } | Select-Object -First 1
        $script:CachedRemediationRulesForGuidance = @{ NonDefault = $nonDefault; Default = $default }
    }
    $nonDefaultRules = $script:CachedRemediationRulesForGuidance.NonDefault
    $defaultRule = $script:CachedRemediationRulesForGuidance.Default

    # Try to match against non-default rules first (most specific first)
    foreach ($rule in $nonDefaultRules) {
        if ($ProductName -like $rule.Pattern) {
            if ($OutputType -eq 'Word') {
                return $rule.WordText
            } else {
                return $rule.TicketText
            }
        }
    }

    # If no match found, use default rule
    if ($defaultRule) {
        if ($OutputType -eq 'Word') {
            return $defaultRule.WordText
        } else {
            return $defaultRule.TicketText
        }
    }

    # Fallback if no rules exist
    if ($OutputType -eq 'Word') {
        return "This application should be updated to the latest version. If available via ConnectWise Automate/RMM or scripting, deploy updates using the patch management system or scripts. Otherwise, manual updates may be required on affected systems."
    } else {
        return "- Update to latest version`r`n  - Deploy via ConnectWise Automate/RMM or scripting if available`r`n  - Otherwise, manual updates required on affected systems"
    }
}
