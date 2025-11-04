#Requires -Modules Microsoft.PowerShell.Utility
<#
.SYNOPSIS
Processes a vulnerability report Excel file:
- Prompts the user to select an input XLSX file.
- Prompts the user to select an output location and filename for the processed file.
- Autofits columns and rows for most worksheets.
- Consolidates data from specific 'Remediation' sheets (excluding 'Linux Remediations')
  into a new 'Source Data' sheet using direct value transfer.
- Creates a Pivot Table from 'Source Data' on a new sheet placed directly after the 'Company' sheet.
- Applies conditional formatting (Value > 0.075 = Light Red) to the 'Max EPSS Score' column in the Pivot Table.
- Adds a color key next to the Pivot Table.
- Resizes Column A on the Pivot Table sheet to width 50.
- Saves the modified workbook to the selected output location.

.DESCRIPTION
This script automates the initial processing of an Excel-based vulnerability report.
It formats the workbook, aggregates relevant data, sets up a Pivot Table with conditional formatting
on a sheet placed after 'Company', adds a color key, resizes column A, and saves the result.

.NOTES
Version: 2.0.0
Requires: Microsoft Excel installed.
Assumption: Sheets matching patterns in $SourceSheetPatterns have identical headers/columns.

.LINK
[About COM Objects](https://docs.microsoft.com/en-us/powershell/scripting/samples/automating-microsoft-excel)
#>

# --- Configuration ---
$SourceSheetPatterns = @("Remediate within *", "Remediate at *")
$ConsolidatedSheetName = "Source Data"
$PivotSheetName = "Proposed Remediations (all)"
$PivotTableName = "VulnPivotTable"
$SheetToExcludeFormatting = "Company" # Worksheet name to skip auto-fitting AND place Pivot sheet after
$KeyHeader = "Key"
# Color key uses only Excel's default color palette (ColorIndex 1-56)
# Standard palette colors: 1=Black, 2=White, 3=Red, 4=Green, 5=Blue, 6=Yellow, 7=Magenta, 8=Cyan
# Gray shades: 15=Gray-25%, 16=Gray-50%, 17=Gray-80%
$KeyData = @(
    @{Text = "Do not touch"; BgColorIndex = 3; FontColorIndex = 2; Strikethrough = $false} # ColorIndex 3=Red (standard palette), ColorIndex 2=White text
    @{Text = "No action needed - auto updates"; BgColorIndex = 4; FontColorIndex = 2; Strikethrough = $false} # ColorIndex 4=Green (standard palette), ColorIndex 2=White text
    @{Text = "Update or patch"; BgColorIndex = 5; FontColorIndex = 2; Strikethrough = $false} # ColorIndex 5=Blue (standard palette), ColorIndex 2=White text
    @{Text = "Uninstall"; BgColorIndex = 16; FontColorIndex = 1; Strikethrough = $false} # ColorIndex 16=Gray-50% (standard palette), ColorIndex 1=Black text
    @{Text = "Already Remediated"; BgColorIndex = 2; FontColorIndex = 1; Strikethrough = $true} # ColorIndex 2=White (standard palette), ColorIndex 1=Black text, strikethrough
    @{Text = "Configuration change needed and further investigation"; BgColorIndex = 6; FontColorIndex = 1; Strikethrough = $false} # ColorIndex 6=Yellow (standard palette), ColorIndex 1=Black text
)
$ConditionalFormatThreshold = 0.075
# Using ColorIndex 3 (Red) from standard palette instead of 22 for better compatibility
$ConditionalFormatColorIndex = 3 # Red from Excel's default color palette
$PivotColumnAWidth = 50
$ExcelPathLimit = 200 # Path length limit before using temporary files

# --- Constants ---
# Excel Pivot Table function constants
$xlMax = -4136
$xlSum = -4157
$xlAutomatic = -4142

# --- Initialize cached assemblies and types ---
# Cache Windows Forms assembly
$null = [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

# Cache kernel32 type for path conversion
try {
    $script:PathConverter = Add-Type -MemberDefinition @"
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetShortPathName(string path, System.Text.StringBuilder shortPath, int shortPathLength);
"@ -Name "PathConverter" -Namespace "Win32" -PassThru -ErrorAction Stop
} catch {
    # Type already exists, reuse it
    $script:PathConverter = [Win32.PathConverter]
}

# --- Helper Functions ---

# Function to safely release COM objects
function Clear-ComObject {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$ComObject
    )
    
    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {
            Write-Warning "Error releasing COM object: $($_.Exception.Message)"
        }
    }
}

# Function to find a worksheet by name
function Find-Worksheet {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,
        
        [Parameter(Mandatory = $true)]
        [string]$SheetName
    )
    
    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Name -eq $SheetName) {
            return $sheet
        }
        Clear-ComObject $sheet
    }
    return $null
}

# Function to select input file
function Get-InputFilePath {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $openFileDialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx"
    $openFileDialog.Title = "Select the INPUT Vulnerability Report Excel File"
    $result = $openFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    } else {
        Write-Warning "Input file selection cancelled."
        return $null
    }
}

# Function to check if Excel is running
function Test-ExcelRunning {
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    return ($excelProcesses.Count -gt 0)
}

# Function to convert long path to short path
function ConvertTo-ShortPath {
    param([string]$LongPath)
    
    if ([string]::IsNullOrWhiteSpace($LongPath)) {
        return $LongPath
    }
    
    try {
        $shortPath = New-Object System.Text.StringBuilder(260)
        $result = $script:PathConverter::GetShortPathName($LongPath, $shortPath, $shortPath.Capacity)
        
        if ($result -gt 0) {
            return $shortPath.ToString()
        } else {
            Write-Warning "Could not convert path to short path: $LongPath"
            return $LongPath
        }
    } catch {
        Write-Warning "Error converting path to short path: $($_.Exception.Message)"
        return $LongPath
    }
}

# Function to create temporary short path
function New-TemporaryShortPath {
    param([string]$OriginalPath)
    
    try {
        $tempDir = [System.IO.Path]::GetTempPath()
        $fileExtension = [System.IO.Path]::GetExtension($OriginalPath)
        $tempFileName = "temp_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8) + $fileExtension
        $tempFile = [System.IO.Path]::Combine($tempDir, $tempFileName)
        
        Copy-Item -Path $OriginalPath -Destination $tempFile -Force
        Write-Host "Created temporary file: $tempFile (Length: $($tempFile.Length) chars)"
        return $tempFile
    } catch {
        Write-Warning "Error creating temporary file: $($_.Exception.Message)"
        return $OriginalPath
    }
}

# Function to select save location
function Get-SaveFilePath {
    param(
        [string]$InitialFileName
    )
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $saveFileDialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx"
    $saveFileDialog.Title = "Select OUTPUT Location to Save Processed Report"
    $saveFileDialog.FileName = $InitialFileName
    $saveFileDialog.OverwritePrompt = $true
    $result = $saveFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName
    } else {
        Write-Warning "Save location selection cancelled."
        return $null
    }
}

# Function to validate and prepare file path
function Resolve-FilePath {
    param(
        [string]$FilePath,
        [ref]$UseTempFile
    )
    
    # Convert to short path
    $shortPath = ConvertTo-ShortPath -LongPath $FilePath
    
    # Remove \\?\ prefix if it exists
    if ($shortPath.StartsWith("\\?\")) {
        $shortPath = $shortPath.Substring(4)
    }
    
    # Check if path is too long
    if ($FilePath.Length -gt $ExcelPathLimit) {
        Write-Host "Path too long for Excel (over $ExcelPathLimit chars). Using temporary file..."
        $shortPath = New-TemporaryShortPath -OriginalPath $FilePath
        $UseTempFile.Value = $true
    }
    
    return $shortPath
}

# Function to validate file accessibility
function Test-FileAccess {
    param([string]$FilePath)
    
    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        throw "Input file does not exist: $FilePath"
    }
    
    try {
        $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $fileStream.Close()
        $fileStream.Dispose()
    } catch {
        throw "Input file is locked or inaccessible: $($_.Exception.Message)"
    }
}

# Function to auto-fit worksheet columns and rows
function Format-Worksheet {
    param(
        [object]$Worksheet,
        [string]$ExcludeSheetName
    )
    
    if ($null -eq $Worksheet -or -not ($Worksheet -is [__ComObject])) {
        return
    }
    
    if ($Worksheet.Name -eq $ExcludeSheetName) {
        Write-Host "  - Skipping formatting for sheet: $($Worksheet.Name)"
        return
    }
    
    try {
        Write-Host "  - Formatting sheet: $($Worksheet.Name)"
        $Worksheet.UsedRange.Columns.AutoFit() | Out-Null
        $Worksheet.UsedRange.Rows.AutoFit() | Out-Null
    } catch {
        Write-Warning "Could not AutoFit sheet '$($Worksheet.Name)'. It might be empty or protected. Error: $($_.Exception.Message)"
    }
}

# Function to find and collect source sheets
function Get-SourceSheets {
    param(
        [object]$Workbook,
        [string[]]$Patterns
    )
    
    $sourceSheets = @()
    $firstValidSheet = $null
    
    Write-Host "Identifying source sheets (patterns: '$($Patterns -join "', '")')..."
    
    foreach ($pattern in $Patterns) {
        $matchingSheets = @($Workbook.Worksheets | Where-Object { $_.Name -like $pattern -or $_.Name -eq $pattern })
        
        foreach ($sheet in $matchingSheets) {
            try {
                if ($null -ne $sheet.UsedRange -and $sheet.UsedRange.Rows.Count -ge 1) {
                    $sourceSheets += $sheet
                    if ($null -eq $firstValidSheet) {
                        $firstValidSheet = $sheet
                        Write-Host "  - Found first valid sheet: $($firstValidSheet.Name)"
                    }
                } else {
                    Write-Host "  - Skipping unusable sheet: $($sheet.Name)"
                    Clear-ComObject $sheet
                }
            } catch {
                Write-Warning "   Could not evaluate sheet '$($sheet.Name)'. Skipping. Error: $($_.Exception.Message)"
                Clear-ComObject $sheet
            }
        }
    }
    
    return @{
        Sheets = $sourceSheets
        FirstValid = $firstValidSheet
    }
}

# Function to copy headers from source sheet
function Copy-Headers {
    param(
        [object]$SourceSheet,
        [object]$TargetSheet
    )
    
    try {
        $sourceCols = $SourceSheet.UsedRange.Columns.Count
        $headerRange = $SourceSheet.Range("A1", $SourceSheet.Cells.Item(1, $sourceCols))
        $headerValues = $headerRange.Value2
        $targetRange = $TargetSheet.Range("A1", $TargetSheet.Cells.Item(1, $sourceCols))
        $targetRange.Value2 = $headerValues
        Write-Host "   Headers copied successfully."
        return $true
    } catch {
        Write-Warning "   Failed header copy. Error: $($_.Exception.Message)"
        return $false
    } finally {
        Clear-ComObject $headerRange
        Clear-ComObject $targetRange
    }
}

# Function to copy data rows from source sheets
function Copy-DataRows {
    param(
        [object[]]$SourceSheets,
        [object]$TargetSheet,
        [int]$StartRow
    )
    
    $destRow = $StartRow
    
    Write-Host "Copying data rows via direct value transfer..."
    
    foreach ($sourceSheet in $SourceSheets) {
        Write-Host "  - Copying from: $($sourceSheet.Name)"
        
        $sourceRange = $null
        $dataRange = $null
        $targetRange = $null
        
        try {
            $sourceRange = $sourceSheet.UsedRange
            $sourceRows = $sourceRange.Rows.Count
            
            if ($sourceRows -gt 1) {
                $sourceCols = $sourceRange.Columns.Count
                $dataRange = $sourceSheet.Range("A2", $sourceSheet.Cells.Item($sourceRows, $sourceCols))
                $dataValues = $dataRange.Value2
                $targetRowCount = $dataRange.Rows.Count
                $targetRange = $TargetSheet.Range($TargetSheet.Cells.Item($destRow, 1), $TargetSheet.Cells.Item($destRow + $targetRowCount - 1, $sourceCols))
                $targetRange.Value2 = $dataValues
                Write-Host "     Data rows copied ($targetRowCount)."
                $destRow += $targetRowCount
            } else {
                Write-Host "   No data rows found."
            }
        } catch {
            Write-Warning "   Failed data row copy. Error: $($_.Exception.Message)"
        } finally {
            Clear-ComObject $dataRange
            Clear-ComObject $targetRange
            Clear-ComObject $sourceRange
        }
    }
    
    return $destRow
}

# --- Main Script ---

$inputFilePath = Get-InputFilePath
if (-not $inputFilePath) {
    Write-Error "No input file selected. Script terminating."
    exit 1
}

$suggestedOutputName = "{0}_Processed.xlsx" -f [System.IO.Path]::GetFileNameWithoutExtension($inputFilePath)
$outputFilePath = Get-SaveFilePath -InitialFileName $suggestedOutputName
if (-not $outputFilePath) {
    Write-Error "No output file location selected. Script terminating."
    exit 1
}

Write-Host "Starting Excel automation..."
Write-Host "Input File: $inputFilePath"
Write-Host "Output File: $outputFilePath"

$excel = $null
$workbook = $null
$sourceDataSheet = $null
$pivotSheet = $null

try {
    # Check if Excel is already running
    Write-Host "Checking for running Excel processes..."
    if (Test-ExcelRunning) {
        Write-Warning "Excel is already running. This may cause COM object conflicts."
        Write-Warning "Consider closing Excel before running this script for best results."
    }
    
    # Create Excel COM Object
    Write-Host "Creating Excel COM object..."
    $excel = New-Object -ComObject Excel.Application
    if ($null -eq $excel) {
        throw "Failed to create Excel COM object. Make sure Microsoft Excel is installed."
    }
    
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # Prepare file paths
    Write-Host "Converting file paths to short format..."
    $useTempFile = $false
    $useTempOutputFile = $false
    
    $shortInputPath = Resolve-FilePath -FilePath $inputFilePath -UseTempFile ([ref]$useTempFile)
    $shortOutputPath = Resolve-FilePath -FilePath $outputFilePath -UseTempFile ([ref]$useTempOutputFile)
    
    Write-Host "Original input path: $inputFilePath (Length: $($inputFilePath.Length) chars)"
    Write-Host "Working input path: $shortInputPath (Length: $($shortInputPath.Length) chars)"
    Write-Host "Original output path: $outputFilePath (Length: $($outputFilePath.Length) chars)"
    Write-Host "Working output path: $shortOutputPath (Length: $($shortOutputPath.Length) chars)"

    # Validate input file
    Write-Host "Validating input file..."
    Test-FileAccess -FilePath $shortInputPath

    # Open Input Workbook
    Write-Host "Opening input workbook..."
    try {
        $workbook = $excel.Workbooks.Open($shortInputPath)
        if ($null -eq $workbook) {
            throw "Failed to open workbook. File may be corrupted or in use."
        }
        Write-Host "Workbook opened successfully."
    } catch {
        throw "Failed to open workbook '$shortInputPath': $($_.Exception.Message)"
    }

    # --- 1. Autofit Columns and Rows ---
    Write-Host "Auto-fitting columns and rows for relevant sheets..."
    foreach ($worksheet in $workbook.Worksheets) {
        Format-Worksheet -Worksheet $worksheet -ExcludeSheetName $SheetToExcludeFormatting
        Clear-ComObject $worksheet
    }

    # --- 2. Consolidate Data ---
    Write-Host "Consolidating data from source sheets (assuming consistent headers, excluding Linux)..."
    
    # Delete existing consolidated sheet if present
    $existingSheet = Find-Worksheet -Workbook $workbook -SheetName $ConsolidatedSheetName
    if ($null -ne $existingSheet) {
        Write-Warning "Deleting existing '$ConsolidatedSheetName' sheet."
        $existingSheet.Delete()
        Clear-ComObject $existingSheet
    } else {
        Write-Host "Sheet '$ConsolidatedSheetName' does not exist."
    }
    
    # Create new consolidated sheet
    Write-Host "Adding new sheet '$ConsolidatedSheetName'..."
    try {
        $lastSheet = $workbook.Worksheets[$workbook.Worksheets.Count]
        $sourceDataSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
        $sourceDataSheet.Name = $ConsolidatedSheetName
        Write-Host "Successfully added sheet '$ConsolidatedSheetName' after '$($lastSheet.Name)'."
        Clear-ComObject $lastSheet
    } catch {
        Write-Warning "Could not add sheet after last. Adding default. Error: $($_.Exception.Message)"
        $sourceDataSheet = $workbook.Worksheets.Add()
        $sourceDataSheet.Name = $ConsolidatedSheetName
        Write-Host "Successfully added sheet '$ConsolidatedSheetName'."
    }
    
    if ($null -eq $sourceDataSheet) {
        Write-Error "Failed '$ConsolidatedSheetName' sheet creation."
        throw "$ConsolidatedSheetName not created."
    }
    
    # Get source sheets
    $sourceSheetInfo = Get-SourceSheets -Workbook $workbook -Patterns $SourceSheetPatterns
    $allSourceSheets = $sourceSheetInfo.Sheets
    $firstValidSheet = $sourceSheetInfo.FirstValid
    
    # Copy headers
    $headersCopied = $false
    if ($null -ne $firstValidSheet) {
        Write-Host "Copying headers from '$($firstValidSheet.Name)'..."
        $headersCopied = Copy-Headers -SourceSheet $firstValidSheet -TargetSheet $sourceDataSheet
    } else {
        Write-Warning "No valid source sheets found for headers."
    }
    
    # Copy data rows
    if ($headersCopied) {
        $destRow = Copy-DataRows -SourceSheets $allSourceSheets -TargetSheet $sourceDataSheet -StartRow 2
    }
    
    # Release source sheet references
    Write-Host "Releasing source sheet COM objects..."
    foreach ($sheet in $allSourceSheets) {
        Clear-ComObject $sheet
    }
    Clear-ComObject $firstValidSheet

    # --- 3. Create Pivot Table ---
    Write-Host "Creating Pivot Table..."
    
    # Delete existing pivot sheet if present
    $existingPivotSheet = Find-Worksheet -Workbook $workbook -SheetName $PivotSheetName
    if ($null -ne $existingPivotSheet) {
        Write-Warning "Deleting existing '$PivotSheetName'."
        $existingPivotSheet.Delete()
        Clear-ComObject $existingPivotSheet
    } else {
        Write-Host "'$PivotSheetName' not found."
    }

    # Find sheet to place pivot after
    Write-Host "Finding '$SheetToExcludeFormatting' sheet to place Pivot Table after..."
    $companySheet = Find-Worksheet -Workbook $workbook -SheetName $SheetToExcludeFormatting
    
    if ($null -ne $companySheet) {
        Write-Host "  - Found '$($companySheet.Name)' sheet."
        Write-Host "Will add Pivot Table sheet after '$($companySheet.Name)'."
        $addSheetAfter = $companySheet
    } else {
        Write-Warning "'$SheetToExcludeFormatting' sheet not found. Pivot Table sheet will be added at the end."
        try {
            $addSheetAfter = $workbook.Worksheets[$workbook.Worksheets.Count]
        } catch {
            $addSheetAfter = $null
        }
    }

    # Add pivot sheet
    try {
        $pivotSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $addSheetAfter)
        $pivotSheet.Name = $PivotSheetName
        $pivotSheet.Tab.ColorIndex = 6 # Yellow
        
        if ($null -ne $addSheetAfter) {
            Write-Host "Successfully added '$PivotSheetName' sheet after '$($addSheetAfter.Name)'."
        } else {
            Write-Host "Successfully added '$PivotSheetName' sheet at the beginning/default position."
        }
    } catch {
        Write-Warning "Could not add pivot sheet after reference. Adding default. Error: $($_.Exception.Message)"
        $pivotSheet = $workbook.Worksheets.Add()
        $pivotSheet.Name = $PivotSheetName
        $pivotSheet.Tab.ColorIndex = 6
    } finally {
        Clear-ComObject $companySheet
        if ($addSheetAfter -ne $companySheet) {
            Clear-ComObject $addSheetAfter
        }
    }
    
    if ($null -eq $pivotSheet) {
        Write-Error "Failed '$PivotSheetName' sheet creation."
        throw "Failed to create pivot sheet"
    }
    
    # Verify source data sheet is still valid
    Write-Host "Verifying `$sourceDataSheet before Pivot Table source assignment..."
    if ($null -eq $sourceDataSheet -or -not ([System.Runtime.InteropServices.Marshal]::IsComObject($sourceDataSheet))) {
        Write-Error "`$sourceDataSheet is NULL or not a valid COM object. This indicates premature release."
        throw "`$sourceDataSheet became invalid before Pivot Table creation."
    }
    
    try {
        $null = $sourceDataSheet.Name
        Write-Host "`$sourceDataSheet appears valid (Name: $($sourceDataSheet.Name))."
    } catch {
        Write-Error "`$sourceDataSheet COM check failed (cannot access .Name). Error: $($_.Exception.Message)"
        throw "`$sourceDataSheet became invalid."
    }

    # Create pivot table
    $pivotSourceRange = $null
    $pivotCache = $null
    $pivotTable = $null
    $dataField1 = $null
    $dataField2 = $null
    
    try {
        Write-Host "Verifying source data..."
        $pivotSourceRange = $sourceDataSheet.UsedRange
        $sourceRowCount = $pivotSourceRange.Rows.Count
        Write-Host "Source Data has $sourceRowCount rows."
        
        if ($sourceRowCount -le 1) {
            Write-Warning "No data rows for Pivot Table."
            throw "No data rows"
        }
        
        Write-Host "Creating Pivot Cache/Table..."
        $pivotCache = $workbook.PivotCaches().Create(1, $pivotSourceRange)
        $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A3"), $PivotTableName)
        Write-Host "Pivot Table object created."
        
        # Configure pivot table fields
        Write-Host "Configuring Pivot Table fields..."
        $rowFieldsToAdd = @("Remediation Type", "Product", "Host Name", "Fix", "IP", "Evidence Path", "Evidence Version")
        
        Write-Host "  - Setting Row Fields..."
        foreach ($fieldName in $rowFieldsToAdd) {
            $ptField = $null
            try {
                Write-Host "    - '$fieldName'"
                $ptField = $pivotTable.PivotFields($fieldName)
                $ptField.Orientation = 1
            } catch {
                Write-Warning "    - Failed '$fieldName'. Error: $($_.Exception.Message)."
            } finally {
                Clear-ComObject $ptField
            }
        }
        
        Write-Host "  - Setting Value Fields..."
        
        # Add EPSS Score field
        $sourceField1 = $null
        try {
            Write-Host "    - Max 'EPSS Score'"
            $sourceField1 = $pivotTable.PivotFields("EPSS Score")
            $dataField1 = $pivotTable.AddDataField($sourceField1)
            $dataField1.Function = $xlMax
            $dataField1.Name = "Max EPSS Score"
        } catch {
            Write-Warning "    - Failed 'EPSS Score'. Error: $($_.Exception.Message)."
        } finally {
            Clear-ComObject $sourceField1
        }
        
        # Add Vulnerability Count field
        $sourceField2 = $null
        try {
            Write-Host "    - Sum 'Vulnerability Count'"
            $sourceField2 = $pivotTable.PivotFields("Vulnerability Count")
            $dataField2 = $pivotTable.AddDataField($sourceField2)
            $dataField2.Function = $xlSum
            $dataField2.Name = "Total Vulnerability Count"
        } catch {
            Write-Warning "    - Failed 'Vulnerability Count'. Error: $($_.Exception.Message)."
        } finally {
            Clear-ComObject $sourceField2
        }
        
        Write-Host "Fields configured."
        
        # Apply conditional formatting
        Write-Host "  - Applying Conditional Formatting..."
        if ($null -ne $dataField1 -and [System.Runtime.InteropServices.Marshal]::IsComObject($dataField1)) {
            try {
                $cfRange = $dataField1.DataRange
                Write-Host "    - Target range: $($cfRange.Address)"
                $cfRange.FormatConditions.Delete()
                $cfConditions = $cfRange.FormatConditions
                $cfCondition = $cfConditions.Add(1, 5, "$ConditionalFormatThreshold")
                $cfCondition.Interior.ColorIndex = $ConditionalFormatColorIndex
                Write-Host "    - Added format (Value > $ConditionalFormatThreshold -> Color $ConditionalFormatColorIndex)."
                
                Clear-ComObject $cfCondition
                Clear-ComObject $cfConditions
                Clear-ComObject $cfRange
            } catch {
                Write-Warning "    - Failed CF. Error: $($_.Exception.Message)"
            }
        } else {
            Write-Warning "    - Cannot apply CF: Max EPSS Score field invalid."
        }
        
        # Add color key
        Write-Host "Adding color key..."
        $headerCell = $null
        $keyColumnObject = $null
        
        try {
            $keyStartCol = $pivotTable.TableRange2.Column + $pivotTable.TableRange2.Columns.Count + 1
            $keyStartRow = $pivotTable.TableRange1.Row
            
            $headerCell = $pivotSheet.Cells.Item($keyStartRow, $keyStartCol)
            $headerCell.Value2 = $KeyHeader
            $headerCell.Font.Bold = $true
            Clear-ComObject $headerCell
            
            $keyCurrentRow = $keyStartRow + 1
            
            foreach ($item in $KeyData) {
                $keyCell = $null
                try {
                    $keyCell = $pivotSheet.Cells.Item($keyCurrentRow, $keyStartCol)
                    $keyCell.Value2 = $item.Text
                    
                    if ($item.ContainsKey('BgColorIndex') -and $null -ne $item.BgColorIndex) {
                        $keyCell.Interior.ColorIndex = $item.BgColorIndex
                    } else {
                        $keyCell.Interior.ColorIndex = $xlAutomatic
                    }
                    
                    if ($item.ContainsKey('FontColorIndex') -and $null -ne $item.FontColorIndex) {
                        $keyCell.Font.ColorIndex = $item.FontColorIndex
                    } else {
                        $keyCell.Font.ColorIndex = 1
                    }
                    
                    # Apply strikethrough formatting if specified
                    if ($item.ContainsKey('Strikethrough')) {
                        $keyCell.Font.Strikethrough = [bool]$item.Strikethrough
                    } else {
                        $keyCell.Font.Strikethrough = $false
                    }
                    
                    Clear-ComObject $keyCell
                    $keyCurrentRow++
                } catch {
                    Write-Warning "Failed to add key item: $($_.Exception.Message)"
                    Clear-ComObject $keyCell
                }
            }
            
            $keyColumnObject = $pivotSheet.Columns.Item($keyStartCol)
            $keyColumnObject.AutoFit() | Out-Null
            Write-Host "Color key added."
        } catch {
            Write-Warning "   - Failed key addition. Error: $($_.Exception.Message)"
        } finally {
            Clear-ComObject $headerCell
            Clear-ComObject $keyColumnObject
        }
        
        # Resize Column A
        Write-Host "Resizing Column A on sheet '$($pivotSheet.Name)' to width $PivotColumnAWidth..."
        $colA = $null
        try {
            $colA = $pivotSheet.Columns("A")
            $colA.ColumnWidth = $PivotColumnAWidth
            Write-Host "  - Column A width set."
        } catch {
            Write-Warning "  - Failed to resize Column A. Error: $($_.Exception.Message)"
        } finally {
            Clear-ComObject $colA
        }

    } catch {
        Write-Error "Pivot Table processing error: $($_.Exception.Message)"
        Write-Error "Stack Trace: $($_.ScriptStackTrace)"
        if ($null -ne $pivotTable -and !$?) {
            Write-Error "Check field names match headers."
        }
    } finally {
        Write-Host "Releasing Pivot Table related COM objects..."
        Clear-ComObject $dataField1
        Clear-ComObject $dataField2
        Clear-ComObject $pivotTable
        Clear-ComObject $pivotCache
        Clear-ComObject $pivotSourceRange
        Write-Host "Pivot Table objects released."
    }

    # --- 4. Save and Close ---
    Write-Host "Preparing to save..."
    $outputDirectory = [System.IO.Path]::GetDirectoryName($shortOutputPath)
    
    if (-not (Test-Path -Path $outputDirectory -PathType Container)) {
        Write-Error "Output directory missing: $outputDirectory"
        throw "Output directory missing."
    } else {
        Write-Host "Output directory verified."
    }
    
    Write-Host "Attempting save to: $shortOutputPath"
    try {
        $workbook.SaveAs($shortOutputPath)
        Write-Host "Workbook saved successfully."
    } catch {
        Write-Error "Failed to SaveAs workbook to '$shortOutputPath'. Error: $($_.Exception.Message)"
        Write-Error "Possible causes: File open, permissions, OneDrive, antivirus."
        throw $_
    }
    
    Write-Host "Closing workbook object..."
    $workbook.Close($false)
    
    # Copy temporary files back if needed
    if ($useTempFile -or $useTempOutputFile) {
        Write-Host "Copying processed file back to original output location..."
        try {
            $finalOutputPath = $outputFilePath
            if ($finalOutputPath.StartsWith("\\?\")) {
                $finalOutputPath = $finalOutputPath.Substring(4)
            }
            
            Copy-Item -Path $shortOutputPath -Destination $finalOutputPath -Force
            Write-Host "File copied to: $finalOutputPath"
            
            # Clean up temporary files
            if ($useTempFile) {
                Remove-Item -Path $shortInputPath -Force -ErrorAction SilentlyContinue
                Write-Host "Temporary input file cleaned up."
            }
            if ($useTempOutputFile) {
                Remove-Item -Path $shortOutputPath -Force -ErrorAction SilentlyContinue
                Write-Host "Temporary output file cleaned up."
            }
        } catch {
            Write-Warning "Failed to copy file to original location: $($_.Exception.Message)"
            Write-Warning "Processed file is available at: $shortOutputPath"
        }
    }
    
    Write-Host "Script finished successfully!"

} catch {
    Write-Error "Unhandled error: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    
    # Provide specific guidance based on error type
    if ($_.Exception.Message -like "*Unable to get the Open property*") {
        Write-Error "Excel COM object issue detected. Possible causes:"
        Write-Error "1. Microsoft Excel is not installed or not properly registered"
        Write-Error "2. Excel is already running and blocking COM access"
        Write-Error "3. File path contains special characters or is too long"
        Write-Error "4. Insufficient permissions to access the file"
    }
    
    if ($null -ne $workbook) {
        try {
            $isSaved = $false
            try {
                $isSaved = $workbook.Saved
            } catch {
                Write-Warning "Check saved status failed."
            }
            
            if (-not $isSaved) {
                Write-Warning "Closing workbook without saving."
            }
            $workbook.Close($false)
        } catch {
            Write-Warning "Close workbook failed: $($_.Exception.Message)"
        }
    }
}
finally {
    Write-Host "Cleaning up COM objects..."
    Clear-ComObject $pivotSheet
    Clear-ComObject $sourceDataSheet
    Clear-ComObject $workbook
    
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        } catch {
            Write-Warning "Excel quit failed."
        }
        Clear-ComObject $excel
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Cleanup complete."
}
