# VScanMagic Alert Processor v3.0

## Overview

The VScanMagic Alert Processor is a PowerShell GUI application designed to process Barracuda XDR security alerts and generate professional, client-facing DOCX reports. It automates the analysis of security alerts by cross-referencing user data from multiple sources and applies the security analysis template for consistent reporting.

## Features

- **Modern GUI Interface**: Easy-to-use Windows Forms interface for file selection and processing
- **Multi-Source Analysis**: Cross-references data from:
  - VScan XLSX files (Barracuda XDR alerts)
  - MFA Status CSV
  - User Security Groups CSV
  - Interactive Sign-Ins CSV
- **Automated Classification**: Automatically classifies alerts as:
  - **True Positive**: Genuine security threats requiring action
  - **False Positive**: Benign events that triggered alerts
  - **Authorized Activity**: Legitimate user behavior (e.g., VPN usage)
- **Professional DOCX Reports**: Generates Word documents with:
  - Service ticket numbers for internal reference
  - Client-friendly language
  - Analysis based on user behavior patterns
  - Actionable recommendations
- **Real-time Logging**: Processing log display for transparency and troubleshooting

## Requirements

- **Operating System**: Windows 10/11 or Windows Server 2016+
- **PowerShell**: 5.1 or later
- **Microsoft Office**: Excel and Word (for COM automation)
- **Permissions**: Script execution policy allowing local scripts

## Installation

1. **Clone or download this repository**:
   ```powershell
   git clone https://github.com/monobrau/vscanmagic.git
   cd vscanmagic
   ```

2. **Set PowerShell execution policy** (if needed):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Verify Microsoft Office is installed**:
   - Excel and Word must be installed and properly licensed
   - Test by running: `New-Object -ComObject Excel.Application`

## Usage

### Launching the Application

Run the PowerShell script:

```powershell
.\VScanAlertProcessor.ps1
```

The GUI will appear with the following components:

### Input Files

1. **VScan XLSX File** (Required)
   - The Barracuda XDR alert export file
   - Expected columns:
     - `TicketNumber`: Service ticket identifier
     - `User`: UserPrincipalName or display name
     - `AlertType`: Type of alert (e.g., "Impossible Travel", "Malicious IP")
     - `IPAddress`: Source IP address
     - `Location`: Geographic location
     - `VPNProvider`: VPN service name (if applicable)

2. **MFA Status CSV** (Optional)
   - User MFA enrollment status
   - Expected columns:
     - `UserPrincipalName` or `DisplayName`
     - `MFAEnabled` or `MFAStatus` (values: "True"/"False" or "Enabled"/"Disabled")

3. **User Security Groups CSV** (Optional)
   - User group memberships
   - Expected columns:
     - `UserPrincipalName` or `DisplayName`
     - `GroupName`: Security group name

4. **Interactive SignIns CSV** (Optional)
   - Historical sign-in data for pattern analysis
   - Expected columns:
     - `UserPrincipalName` or `DisplayName`
     - `IPAddress`: Sign-in IP address
     - `Location`: Geographic location

5. **Output Folder** (Required)
   - Destination folder for generated DOCX reports
   - Reports are named: `SecurityAlertReport_YYYYMMDD_HHMMSS.docx`

### Processing Alerts

1. Select all required and optional input files using the "Browse..." buttons
2. Select the output folder
3. Click "Process Alerts & Generate Reports"
4. Monitor the processing log for real-time status updates
5. Upon completion, a success message will display the report location

## Report Format

Each generated DOCX report contains:

### Structure (Per Alert)

```
Service Ticket: [Ticket Number]

Good afternoon,

[Analysis paragraph including:
 - User identification
 - Alert type description
 - MFA status
 - Administrative role status
 - Behavioral analysis
 - Classification (True Positive/False Positive/Authorized Activity)
 - Recommended actions]

Sincerely,

Chris Knospe
```

### Example Output

```
Service Ticket: INC0012345

Good afternoon,

We have reviewed the security alert for john.doe@company.com regarding Impossible Travel.
The user's account is protected with multi-factor authentication. This user has administrative
privileges in your environment.

After analyzing the sign-in logs and cross-referencing with the user's typical behavior, we
have classified this alert as: Authorized Activity.

This appears to be normal behavior for this user who regularly uses VPN services or accesses
the system from multiple locations.

Sincerely,

Chris Knospe
```

## Analysis Logic

The script performs the following analysis for each alert:

1. **User Identification**: Extracts user from alert data
2. **MFA Check**: Determines if MFA is enabled for the account
3. **Role Assessment**: Identifies if user has administrative privileges
4. **Behavioral Analysis**:
   - **VPN Usage Pattern**: Checks if user frequently uses VPNs (>5 VPN sign-ins)
   - **Location Diversity**: Analyzes if user regularly signs in from multiple locations (>3 unique locations)
5. **Classification**:
   - **Impossible Travel**:
     - VPN user or diverse locations → "Authorized Activity"
     - Otherwise → "True Positive"
   - **Malicious IP**:
     - Frequent VPN user → "Authorized Activity"
     - Otherwise → "True Positive"

## Customization

You can customize the script by modifying these variables at the top of the script:

```powershell
$Script:Config = @{
    AppTitle = "VScanMagic Alert Processor v3.0"
    DefaultAnalyst = "Chris Knospe"           # Change to your name
    GreetingSalutation = "Good afternoon,"     # Customize greeting
    ClosingSalutation = "Sincerely,"           # Customize closing
}
```

### Adding Custom Alert Types

To handle additional alert types, modify the `Get-AlertAnalysis` function:

```powershell
elseif ($alertType -match "YourCustomAlertType") {
    # Add your custom logic here
    $classification = "True Positive"
    $actionRequired = "Your recommended action..."
}
```

## Troubleshooting

### Common Issues

1. **"Cannot create COM object" error**
   - **Cause**: Microsoft Office not installed or not properly registered
   - **Solution**: Verify Excel and Word are installed and licensed

2. **"Access denied" when reading files**
   - **Cause**: Files are open in another application or permissions issue
   - **Solution**: Close all Excel/Word instances and verify file permissions

3. **Missing columns in XLSX/CSV**
   - **Cause**: Input files don't match expected format
   - **Solution**: Verify column headers match the expected format (see Input Files section)

4. **Empty or incorrect analysis**
   - **Cause**: CSV files not provided or data not matching
   - **Solution**: Ensure UserPrincipalName or DisplayName columns exist and contain matching values

### Debug Mode

To enable verbose logging, add the following at the start of the script:

```powershell
$VerbosePreference = 'Continue'
$DebugPreference = 'Continue'
```

## File Format Requirements

### VScan XLSX Expected Format

| TicketNumber | User | AlertType | IPAddress | Location | VPNProvider |
|--------------|------|-----------|-----------|----------|-------------|
| INC0012345 | john.doe@company.com | Impossible Travel | 192.168.1.1 | New York, US | NordVPN |

### MFA Status CSV Expected Format

| UserPrincipalName | MFAEnabled |
|-------------------|------------|
| john.doe@company.com | True |

### User Security Groups CSV Expected Format

| UserPrincipalName | GroupName |
|-------------------|-----------|
| john.doe@company.com | Domain Admins |

### Interactive SignIns CSV Expected Format

| UserPrincipalName | IPAddress | Location |
|-------------------|-----------|----------|
| john.doe@company.com | 192.168.1.1 | New York, US |

## Version History

### v3.0.0 (Current)
- Initial release of Alert Processor with GUI
- Automated DOCX report generation
- Multi-source data analysis
- Behavioral pattern analysis for VPN/location anomalies

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues, questions, or feature requests:
- GitHub Issues: https://github.com/monobrau/vscanmagic/issues
- Email: [Your support email]

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request with a clear description of changes

## Acknowledgments

- Built on PowerShell Windows Forms
- Uses Microsoft Office COM automation
- Designed for MSP security operations teams
