# VScanMagic

A PowerShell automation suite for processing vulnerability scan Excel files and generating professional security assessment reports.

## Versions

- **VScanMagic v2** (`vscanmagicv2.ps1`) - Original command-line script for Excel processing
- **VScanMagic v3 GUI** (`VScanMagic-GUI.ps1`) - Modern GUI application with Word report generation
- **VScanMagic v3 EXE** (`VScanMagic.exe`) - Standalone executable version (no PowerShell required)

## VScanMagic v3 GUI Features

The new GUI version offers enhanced functionality for vulnerability reporting:

- **Graphical User Interface**: User-friendly Windows Forms interface
- **Risk Score Calculation**: Automatic computation of composite risk scores using EPSS and CVSS data
  - Formula: **Risk Score = EPSS Score × Average CVSS**
  - Average CVSS calculated from weighted severity counts (Critical×9.0 + High×7.0 + Medium×5.0 + Low×3.0)
  - Dynamic severity thresholds adapt to your data (percentages of maximum risk score)
- **Word Report Generation**: Professional DOCX reports with:
  - Executive summary with client name and scan date
  - Scoring methodology documentation
  - Color-coded Top 10 vulnerability table with dynamic severity thresholds
  - Dynamic risk score legend (thresholds adapt to your data)
  - Pie chart showing vulnerability distribution by product/system
  - Detailed findings with remediation guidance
  - Affected systems listing
- **Excel Report Generation**: Processed XLSX files with:
  - Auto-fitted columns and rows
  - Consolidated "Source Data" sheet from all remediation sheets
  - Pivot table with vulnerability analysis
  - Conditional formatting for high-risk items (EPSS > 0.075)
  - Color-coded remediation status key
- **Windows Version Consolidation**: Automatically groups Windows versions (Server 2012, Win 10, Win 11)
- **Smart Filtering**: Excludes auto-updating browsers and duplicate entries
- **Remediation Intelligence**: Context-aware guidance for different vulnerability types
- **ConnectWise Automate Integration**: Recommendations aligned with CW Automate patch management
- **User Settings Persistence**: Save and reuse company information, contact details, and prepared-by name
- **Real-time Processing Log**: View progress and status updates within the GUI

### Dynamic Color-Coded Risk Levels

The application uses **dynamic severity thresholds** that adapt to your specific vulnerability data. Thresholds are calculated as percentages of the maximum risk score found in your dataset, ensuring appropriate color coding regardless of the data range.

**Five Risk Levels** (from highest to lowest):
- **Crimson Red (Critical)**: Risk Score ≥ 100% of maximum (top tier vulnerabilities)
- **Orange-Red (Very High)**: Risk Score ≥ 70% of maximum
- **Dark Orange (High)**: Risk Score ≥ 50% of maximum
- **Orange (Medium-High)**: Risk Score ≥ 30% of maximum
- **Yellow (Medium)**: Risk Score ≥ 0% (baseline - all items need attention)

**Note**: The color gradient uses yellow-to-red (no greens) to emphasize that all vulnerabilities in the Top 10 require attention, with severity indicated by the intensity of the color.

## VScanMagic v2 Features

- **Automatic Formatting**: Auto-fits columns and rows for all worksheets (except specified exclusions)
- **Data Consolidation**: Consolidates data from multiple "Remediation" sheets into a single "Source Data" sheet
- **Pivot Table Creation**: Automatically creates a pivot table with configured fields and conditional formatting
- **EPSS Score Highlighting**: Applies conditional formatting to highlight high-risk vulnerabilities (EPSS Score > 0.075)
- **Color Key**: Adds a color-coded key for remediation status tracking
- **Path Length Handling**: Automatically handles long file paths using temporary files when necessary
- **COM Object Management**: Proper cleanup of Excel COM objects to prevent memory leaks

## Requirements

### For VScanMagic v3 GUI
- Windows PowerShell 5.1 or later
- Microsoft Excel installed and properly registered
- Microsoft Word installed (for DOCX report generation)
- .NET Framework (included with Windows)
- PowerShell execution policy that allows script execution

### For VScanMagic v2
- Windows PowerShell 5.1 or later
- Microsoft Excel installed and properly registered
- PowerShell execution policy that allows script execution

## Installation

### Option 1: Standalone Executable (Recommended)

1. Download `VScanMagic.zip` from the repository
2. Extract `VScanMagic.exe` to your desired location
3. Double-click to run - no PowerShell installation required!

**Note:** Microsoft Excel and Word must still be installed on the system for the application to function.

### Option 2: PowerShell Script

1. Clone this repository:
   ```bash
   git clone https://github.com/monobrau/vscanmagic.git
   cd vscanmagic
   ```
   Or download the scripts directly:
   - `VScanMagic-GUI.ps1` (recommended for most users)
   - `vscanmagicv2.ps1` (command-line version)

2. Ensure Microsoft Excel and Word are installed on your system (for v3 GUI)

3. If needed, adjust your PowerShell execution policy:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

### Building the Executable

To build your own executable from the PowerShell script:

1. Install the `ps2exe` PowerShell module:
   ```powershell
   Install-Module -Name ps2exe -Scope CurrentUser
   ```

2. Run the build script:
   ```powershell
   .\BuildExeFinal.ps1
   ```

This will create `VScanMagic.exe` with the custom icon (`VScanMagic.ico`) and package it in `VScanMagic.zip`.

## Usage

### VScanMagic v3 (EXE or GUI Script)

**Using the Executable:**
1. Double-click `VScanMagic.exe` to launch the application

**Using the PowerShell Script:**
1. Run the GUI script:
   ```powershell
   .\VScanMagic-GUI.ps1
   ```

2. In the GUI window:
   - Click **Settings** (top right) to configure your company information, contact details, and "Prepared By" name (saved for future use)
   - **Browse** for your input XLSX vulnerability scan file
   - Enter the **Client Name** (auto-populated from filename if detected)
   - Select the **Scan Date**
   - Choose output options (Excel Report and/or Word Report)
   - Select the **Output Directory**
   - Click **Generate**

3. The application will:
   - **Auto-detect** and consolidate all "Remediate within *" and "Remediate at *" sheets
   - Exclude "Company" and "Linux Remediations" sheets
   - Read and analyze vulnerability data
   - Calculate composite risk scores (EPSS × Average CVSS)
   - Calculate dynamic severity thresholds based on maximum risk score
   - Identify top 10 vulnerabilities by risk score
   - Generate professional Word report with:
     - Dynamic color-coded Top 10 table (thresholds adapt to your data)
     - Risk score legend showing percentage-based thresholds
     - Pie chart visualizing vulnerability distribution
     - Detailed findings with remediation guidance
     - Affected systems listing
   - Generate processed Excel file with pivot tables and conditional formatting (if selected)
   - Provide detailed remediation guidance

4. View the processing log in real-time within the GUI

5. After generation, use the **Open** buttons to view your reports directly from the application

### Required Excel Structure

**Auto-Sheet Detection:**
VScanMagic v3 automatically finds and consolidates data from multiple sheets:
- **Included**: All sheets matching "Remediate within *" or "Remediate at *" patterns
- **Excluded**: "Company", "Linux Remediations"
- No manual sheet selection needed - works just like v2!

**Flexible Column Detection:**
The script uses intelligent column mapping and can detect variations of common column names:

**Minimum Required:**
- **Product/Software** column (required) - Variations recognized:
  - Product, Software, Application, App, Program, Title, Product Name, Software Name

**Optional Columns** (with multiple name variations recognized):
- **Host Name**: Host Name, Hostname, Computer, Computer Name, Device, System, Machine
- **IP Address**: IP, IP Address, IPAddress, Address
- **Critical**: Critical, Crit, Critical Count, Critical Vulnerabilities
- **High**: High, High Count, High Vulnerabilities
- **Medium**: Medium, Med, Medium Count, Medium Vulnerabilities
- **Low**: Low, Low Count, Low Vulnerabilities
- **Vulnerability Count**: Vulnerability Count, Vuln Count, Total Vulnerabilities, Total Vulns, Count
- **EPSS Score**: EPSS Score, EPSS, Exploit Prediction Score, Max EPSS Score

**Smart Features:**
- Automatic multi-sheet consolidation (like v2)
- Case-insensitive column matching
- Partial column name matching
- Automatic vulnerability count calculation if not provided (sums Critical + High + Medium + Low)
- Handles missing columns gracefully (uses defaults)
- Filters out rows with no product name or zero vulnerabilities
- Progress logging for large datasets

### VScanMagic v2 (Command Line)

1. Run the script:
   ```powershell
   .\vscanmagicv2.ps1
   ```

2. When prompted, select your input Excel file (`.xlsx` format)

3. When prompted, choose the output location and filename for the processed file

4. The script will:
   - Process and format all worksheets
   - Create a consolidated "Source Data" sheet
   - Generate a pivot table on a new "Proposed Remediations (all)" sheet
   - Apply conditional formatting and add a color key
   - Save the processed workbook

## Configuration

The script can be customized by modifying the configuration variables at the top of the script:

- `$SourceSheetPatterns`: Patterns to match source sheets for consolidation (default: `"Remediate within *"`, `"Remediate at *"`)
- `$ConsolidatedSheetName`: Name for the consolidated data sheet (default: `"Source Data"`)
- `$PivotSheetName`: Name for the pivot table sheet (default: `"Proposed Remediations (all)"`)
- `$SheetToExcludeFormatting`: Sheet name to skip auto-fitting and place pivot sheet after (default: `"Company"`)
- `$ConditionalFormatThreshold`: EPSS score threshold for conditional formatting (default: `0.075`)
- `$PivotColumnAWidth`: Width for column A in the pivot table sheet (default: `50`)

## Pivot Table Structure

The pivot table includes the following fields:

### Row Fields:
- Remediation Type
- Product
- Host Name
- Fix
- IP
- Evidence Path
- Evidence Version

### Value Fields:
- Max EPSS Score (with conditional formatting)
- Total Vulnerability Count

## Color Key

The script adds a color-coded key indicating remediation status:
- **Red**: Do not touch
- **Green**: No action needed - auto updates
- **Blue**: Update or patch
- **Gray**: Uninstall
- **White (Strikethrough)**: Already Remediated
- **Yellow**: Configuration change needed and further investigation

## Troubleshooting

### Excel COM Object Errors
- Ensure Excel is not already running when you start the script
- Verify Microsoft Excel is properly installed and registered
- Check that file paths don't contain special characters or exceed length limits

### File Access Issues
- Ensure the input file is not open in Excel or another program
- Check file permissions for both input and output locations
- If using OneDrive, ensure files are synced locally

### Path Length Issues
- The script automatically handles paths longer than 200 characters using temporary files
- If you encounter path issues, try using shorter directory names or moving files closer to the root

## Risk Score Calculation Details

### Formula
```
Risk Score = EPSS Score × Average CVSS
```

Where:
- **EPSS Score**: Exploit Prediction Scoring System score (0.0 to 1.0)
- **Average CVSS**: Weighted average CVSS equivalent score calculated as:
  ```
  (Critical × 9.0 + High × 7.0 + Medium × 5.0 + Low × 3.0) / Total Vulnerabilities
  ```

### Dynamic Thresholds
The application calculates severity thresholds dynamically based on the maximum risk score in your dataset:
- Finds the highest risk score in the Top 10 vulnerabilities
- Calculates thresholds as percentages: 100%, 70%, 50%, 30%, 0%
- Ensures appropriate color coding regardless of data range
- Displays thresholds in the Word report legend with actual values and percentages

**Example**: If your maximum risk score is 8.5:
- Critical: ≥ 8.50 (100%)
- Very High: ≥ 5.95 (70%)
- High: ≥ 4.25 (50%)
- Medium-High: ≥ 2.55 (30%)
- Medium: ≥ 0.00 (0%)

## Version History

- **v3.0.0**: 
  - Added standalone executable version (VScanMagic.exe)
  - Implemented dynamic severity thresholds
  - Added pie chart visualization
  - Enhanced path handling for EXE execution
  - Added user settings persistence
  - Improved risk score calculation and documentation
- **v1.20.0**: Refined COM object release logic for Company sheet search and Pivot sheet placement
- Previous versions: Initial functionality and bug fixes

## License

This script is provided as-is for use in vulnerability reporting workflows.

**Copyright (c) 2025 Chris Knospe**

## Contributing

Contributions, issues, and feature requests are welcome! Please feel free to submit pull requests or open issues for any problems or suggestions.

## Disclaimer

This script automates Excel operations using COM objects. Ensure you have proper backups of your data before running the script. The authors are not responsible for any data loss or corruption that may occur during script execution.

