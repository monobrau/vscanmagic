# VScanMagic

A PowerShell tool for vulnerability management and report generation from ConnectSecure.

## Tools Included

### VScanMagic v4 - Vulnerability Report Generator (`VScanMagic-GUI.ps1`)

A GUI application that processes vulnerability scan reports from **ConnectSecure** and generates professional Word and Excel reports with dynamic severity thresholds, email templates, ticket notes, and time estimates.

**Features:**
- **Download Standard Reports** – Download all 5 ConnectSecure reports (All Vulnerabilities, Suppressed, External Scan, Executive Summary DOCX, Pending EPSS) with retry logic
- **Professional Word Reports** – Color-coded Top Ten vulnerabilities reports with dynamic severity thresholds
- **Excel Spreadsheets** – Detailed vulnerability spreadsheets with risk scoring
- **Email Templates** – Professional email templates for client communication
- **Ticket Notes** – ConnectWise-compatible ticket notes with randomized content
- **Time Estimates** – RMIT/RMIT+ time estimation with covered software support
- **Dynamic Severity Thresholds** – Adapts risk score thresholds based on your data

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

#### ConnectSecure Download Flow

1. Configure **API Settings** (Base URL, Tenant, Client ID, Client Secret)
2. Select **Company** and **Scan Date**
3. Check desired reports (All Vulnerabilities, Executive Summary, Suppressed, Pending EPSS, External Scan)
4. Click **Download Standard Reports Only** – downloads with retry logic and progress
5. Or click **Generate** to download and process in one step

#### File Processing Flow

1. Use **Browse** to select a previously downloaded Pending EPSS report (XLSX)
2. Enter **Client** name and **Output Directory**
3. Select **Output Options** (Word, Excel, Email Template, Ticket Instructions, Time Estimate)
4. Click **Generate** to create reports

### Input File Requirements

VScanMagic requires the **Pending EPSS report** from ConnectSecure. The Excel file must include columns such as Product/Software, EPSS Score, CVSS Severity, Host Name, IP Address, and vulnerability details.

### Documentation

- [ConnectSecure API Usage](ConnectSecure-API-Usage.md) – API configuration and usage
- [Quick Start](QUICK_START.md) – Step-by-step guide
- [Release Notes](RELEASE_v4.0.0.md) – v4.0.0 changelog

---

## Contributing

Contributions, issues, and feature requests are welcome. Please submit pull requests or open issues for problems or suggestions.

## License

This project is provided as-is for use in vulnerability reporting workflows. See [LICENSE](LICENSE) for details.

## Disclaimer

These scripts automate Office operations using COM objects. Ensure you have proper backups before running. The authors are not responsible for any data loss or corruption during script execution.
