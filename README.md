# VScanMagic

A PowerShell tool for vulnerability management and report generation from ConnectSecure.

## Tools Included

### VScanMagic v4 - Vulnerability Report Generator (`VScanMagic-GUI.ps1`)

A GUI application that processes vulnerability scan reports from **ConnectSecure** and generates professional Word and Excel reports with dynamic severity thresholds, email templates, ticket notes, and time estimates.

**Features:**
- **Download Standard Reports** – Download the 5 core ConnectSecure reports (All Vulnerabilities, Suppressed, External Scan, Executive Summary DOCX, Pending EPSS) with retry logic
- **Download Custom Report** – Choose any standard report from ConnectSecure (company-scoped) with format options (XLSX, DOCX, PDF)
- **Download Global Reports** – Download global reports (tenant-wide, no company filter) in their own dialog with format options
- **Professional Word Reports** – Color-coded Top Ten vulnerabilities reports with dynamic severity thresholds
- **Excel Spreadsheets** – Detailed vulnerability spreadsheets with risk scoring
- **Email Templates** – Professional email templates for client communication
- **Ticket Notes** – ConnectWise-compatible ticket notes with randomized content
- **Time Estimates** – RMIT/RMIT+ time estimation with covered software support
- **Dynamic Severity Thresholds** – Adapts risk score thresholds based on your data
- **General Help & API Help** – In-app help for workflow guidance and API credential setup

**Recent updates (4.0.2):** Structured output paths with Network Documentation/Vulnerability Scans; quarter folder format (2026 - Q1); Processing Summary and Report Folder History for quick access to output folders; bulk processing skips completion popups; removed View Generated Reports section; Company Folder Mappings auto-normalization.

---

## VScanMagic v4 - Quick Start

### Requirements

- **Windows PowerShell 5.1 or later**
- **Microsoft Excel** and **Microsoft Word** installed and properly registered
- **PowerShell execution policy** that allows script execution

### Installation

#### Option 1: Standalone Executable (Recommended)

1. Download `VScanMagic.zip` from the [latest release](https://github.com/monobrau/vscanmagic/releases)
2. Extract the ZIP – keep `VScanMagic.exe`, `VScanMagic-Modules`, and `ConnectSecure-API.ps1` in the same folder
3. Double-click `VScanMagic.exe` to launch

#### Option 2: PowerShell Script

1. Clone this repository
2. Run:
   ```powershell
   .\VScanMagic-GUI.ps1
   ```

### Usage

**Main button:** The green **Download and Generate Reports** button is the primary action. It both downloads reports from ConnectSecure (when API is configured) and processes data to generate Word, Excel, and other outputs. Use it as your main workflow step.

#### ConnectSecure Download Flow

1. Configure **API Settings** (Base URL, Tenant, Client ID, Client Secret) – use **API Help** for credential setup
2. Select **Company** and **Scan Date**
3. Check desired reports (All Vulnerabilities, Executive Summary, Suppressed, Pending EPSS, External Scan)
4. Click **Download Standard Reports Only** – downloads the 5 core reports with retry logic
5. Click **Download Custom…** to pick any standard report with format (XLSX/DOCX/PDF)
6. Click **Download Global Reports** for tenant-wide reports (no company filter)
7. Or click the main **Download and Generate Reports** button to download and process in one step

#### File Processing Flow

1. Use **Browse** to select a previously downloaded **All Vulnerabilities** report (XLSX)
2. Enter **Client** name and **Output Directory**
3. Select **Output Options** (Word, Excel, Email Template, Ticket Instructions, Time Estimate)
4. Click the main **Download and Generate Reports** button – it downloads from ConnectSecure (if configured) and processes files to create all selected outputs

### Input File Requirements

VScanMagic uses the **All Vulnerabilities report** from ConnectSecure as the primary input. The Excel file must include columns such as Software Name, Host Name, IP, EPSS Score, Severity, and vulnerability details. The full vulnerability list format (Critical/High/Medium/Low/END OF LIFE sheets) is supported.

### Documentation

- [ConnectSecure API Usage](ConnectSecure-API-Usage.md) – API configuration and usage
- [Quick Start](QUICK_START.md) – Step-by-step guide
- [Release Notes](https://github.com/monobrau/vscanmagic/releases) – Changelog and downloads

---

## Contributing

Contributions, issues, and feature requests are welcome. Please submit pull requests or open issues for problems or suggestions.

## License

This project is provided as-is for use in vulnerability reporting workflows. See [LICENSE](LICENSE) for details.

## Disclaimer

These scripts automate Office operations using COM objects. Ensure you have proper backups before running. The authors are not responsible for any data loss or corruption during script execution.
