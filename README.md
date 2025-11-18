# VScanMagic

A suite of PowerShell tools for vulnerability management and security alert processing.

## Tools Included

### 1. VScanMagic v2 - Vulnerability Report Processor (`vscanmagicv2.ps1`)
Automates the processing of vulnerability report Excel files, performing data consolidation, formatting, and pivot table creation with conditional formatting.

### 2. VScanMagic v3 - Vulnerability Report Generator (`VScanMagic-GUI.ps1`) ⭐ NEW
A GUI application that processes vulnerability scan Excel files and generates professional Word and Excel reports with dynamic severity thresholds, email templates, and ticket notes.

### 3. VScanMagic Alert Processor v3 - Security Alert Analysis Tool (`VScanAlertProcessor.ps1`) ⭐ NEW
A GUI application that processes Barracuda XDR security alerts and generates professional, client-facing DOCX reports with automated threat classification.

---

## VScanMagic v2 - Vulnerability Report Processor

A PowerShell script that automates the processing of vulnerability report Excel files.

## Features

- **Automatic Formatting**: Auto-fits columns and rows for all worksheets (except specified exclusions)
- **Data Consolidation**: Consolidates data from multiple "Remediation" sheets into a single "Source Data" sheet
- **Pivot Table Creation**: Automatically creates a pivot table with configured fields and conditional formatting
- **EPSS Score Highlighting**: Applies conditional formatting to highlight high-risk vulnerabilities (EPSS Score > 0.075)
- **Color Key**: Adds a color-coded key for remediation status tracking
- **Path Length Handling**: Automatically handles long file paths using temporary files when necessary
- **COM Object Management**: Proper cleanup of Excel COM objects to prevent memory leaks

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Excel installed and properly registered
- PowerShell execution policy that allows script execution

## Installation

1. Clone this repository or download the `vscanmagicv2.ps1` file
2. Ensure Microsoft Excel is installed on your system
3. If needed, adjust your PowerShell execution policy:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

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

## Version History

- **v1.20.0**: Refined COM object release logic for Company sheet search and Pivot sheet placement
- Previous versions: Initial functionality and bug fixes

## License

This script is provided as-is for use in vulnerability reporting workflows.

## VScanMagic v3 - Vulnerability Report Generator

A GUI application that processes vulnerability scan Excel files and generates professional Word and Excel reports with dynamic severity thresholds, email templates, and ticket notes.

### Features

- **Professional Word Reports**: Generates color-coded Top Ten vulnerabilities reports with dynamic severity thresholds
- **Excel Spreadsheets**: Creates detailed vulnerability spreadsheets with risk scoring
- **Email Templates**: Generates professional email templates for client communication
- **Ticket Notes**: Creates ConnectWise-compatible ticket notes with randomized content
- **Dynamic Severity Thresholds**: Automatically adapts risk score thresholds based on your data
- **Risk Score Calculation**: Uses EPSS scores and CVSS equivalents for accurate risk assessment

### Requirements

- **Windows PowerShell 5.1 or later**
- **Microsoft Excel** installed and properly registered
- **Microsoft Word** installed and properly registered
- **PowerShell execution policy** that allows script execution
- **Input File**: **Pending EPSS report** exported from the **ConnectSecure portal** (Excel `.xlsx` format)

### Installation

#### Option 1: Standalone Executable (Recommended)

1. Download `VScanMagic.zip` from the latest release
2. Extract `VScanMagic.exe` to your desired location
3. Double-click `VScanMagic.exe` to launch

#### Option 2: PowerShell Script

1. Clone this repository or download `VScanMagic-GUI.ps1`
2. Ensure Microsoft Excel and Word are installed
3. If needed, adjust your PowerShell execution policy:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
4. Run the script:
   ```powershell
   .\VScanMagic-GUI.ps1
   ```

### Usage

1. **Export the Pending EPSS report** from the ConnectSecure portal
2. Launch VScanMagic v3 (run `VScanMagic.exe` or `VScanMagic-GUI.ps1`)
3. **Select the Pending EPSS report** Excel file using the "Browse..." button
4. Enter the **Client Name** and **Output Directory**
5. Choose your **Output Options**:
   - Generate Word Report (Top Ten vulnerabilities)
   - Generate Excel Spreadsheet
   - Generate Email Template
   - Generate Ticket Instructions
   - Generate Ticket Notes (ConnectWise format)
6. Click **Generate** to create the reports
7. Use the **Open** buttons to view generated files

### Input File Requirements

**Important**: VScanMagic v3 requires the **Pending EPSS report** exported from the **ConnectSecure portal**.

The Excel file must contain the following columns:
- Product/Software
- EPSS Score
- CVSS Severity (Critical, High, Medium, Low)
- Host Name
- IP Address
- Additional vulnerability details

### Dynamic Severity Thresholds

The application uses **dynamic severity thresholds** that adapt to your specific vulnerability data. Thresholds are calculated as percentages of the maximum risk score found in your dataset:

- **Crimson Red (Critical)**: Risk Score ≥ 100% of maximum
- **Orange-Red (Very High)**: Risk Score ≥ 70% of maximum
- **Dark Orange (High)**: Risk Score ≥ 50% of maximum
- **Orange (Medium-High)**: Risk Score ≥ 30% of maximum
- **Yellow (Medium)**: Risk Score ≥ 0% (baseline)

### Risk Score Calculation

Risk Score = EPSS Score × Average CVSS

Where Average CVSS is calculated as:
```
(Critical × 9.0 + High × 7.0 + Medium × 5.0 + Low × 3.0) / Total Vulnerabilities
```

### Ticket Notes Feature

Generate ConnectWise-compatible ticket notes with:
- Randomized task descriptions (5-8 variations)
- Randomized steps performed (11 categories with 4 variations each)
- Professional formatting with bold markdown headers
- One-click generation and clipboard copy

### Troubleshooting

#### Excel/Word COM Object Errors
- Ensure Excel and Word are not already running when you start the application
- Verify Microsoft Office is properly installed and licensed
- Check that file paths don't contain special characters

#### File Access Issues
- Ensure the input Excel file is not open in Excel or another program
- Check file permissions for both input and output locations
- If using OneDrive, ensure files are synced locally

#### Missing Columns Error
- Verify you're using the **Pending EPSS report** from ConnectSecure portal
- Ensure the Excel file contains required columns (Product, EPSS Score, CVSS Severity, etc.)
- Check that the file hasn't been modified or corrupted

## VScanMagic Alert Processor v3 - Security Alert Analysis Tool

**For detailed documentation on the Alert Processor, see [ALERT_PROCESSOR_README.md](ALERT_PROCESSOR_README.md)**

### Quick Start

The Alert Processor is a GUI application for analyzing Barracuda XDR security alerts:

```powershell
.\VScanAlertProcessor.ps1
```

### Key Features
- **GUI Interface**: Easy-to-use Windows Forms application
- **Multi-Source Analysis**: Cross-references VScan XLSX with MFA status, security groups, and sign-in logs
- **Automated Classification**: Classifies alerts as True Positive, False Positive, or Authorized Activity
- **DOCX Report Generation**: Creates professional, client-facing Word documents
- **Behavioral Analysis**: Analyzes VPN usage patterns and location diversity

### Required Inputs
- VScan XLSX file (Barracuda XDR alerts)
- MFA Status CSV (optional)
- User Security Groups CSV (optional)
- Interactive Sign-Ins CSV (optional)
- Output folder for reports

### Report Template

The tool uses this analysis framework for each alert:

1. **Investigation**: Identifies user, alert type, and observables
2. **Cross-Reference**: Checks MFA status and administrative roles
3. **Behavioral Analysis**: Analyzes VPN/location patterns from sign-in logs
4. **Classification**: Determines True Positive, False Positive, or Authorized Activity
5. **Recommendations**: Provides clear, actionable guidance for clients

Each report includes:
- Service ticket number (for internal reference)
- Client-friendly greeting and closing
- Plain language analysis (no technical jargon)
- Clear classification and action items

---

## Contributing

Contributions, issues, and feature requests are welcome! Please feel free to submit pull requests or open issues for any problems or suggestions.

## Disclaimer

These scripts automate Office operations using COM objects. Ensure you have proper backups of your data before running the scripts. The authors are not responsible for any data loss or corruption that may occur during script execution.

