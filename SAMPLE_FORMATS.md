# Sample File Formats for VScanMagic Alert Processor

This document provides examples of the expected file formats for the VScanMagic Alert Processor.

## VScan XLSX File Format

The main alert file should be an Excel (.xlsx) file with the following columns:

### Required Columns

| Column Name | Description | Example |
|-------------|-------------|---------|
| TicketNumber | Service ticket identifier | INC0012345 |
| User | User's email or display name | john.doe@company.com |
| AlertType | Type of security alert | Impossible Travel |
| IPAddress | Source IP address from the alert | 192.168.1.100 |
| Location | Geographic location | New York, NY, USA |
| VPNProvider | VPN service name (if applicable) | NordVPN |

### Sample Data

```
TicketNumber    User                    AlertType           IPAddress       Location            VPNProvider
INC0012345      john.doe@company.com    Impossible Travel   192.168.1.100   New York, NY, USA   NordVPN
INC0012346      jane.smith@company.com  Malicious IP        10.0.0.50       London, UK          ExpressVPN
INC0012347      bob.jones@company.com   Impossible Travel   172.16.0.10     Tokyo, Japan
```

---

## MFA Status CSV Format

This file contains MFA enrollment status for users.

### Required Columns

| Column Name | Description | Example |
|-------------|-------------|---------|
| UserPrincipalName | User's email address | john.doe@company.com |
| DisplayName | User's display name | John Doe |
| MFAEnabled | MFA status (True/False or Enabled/Disabled) | True |

**Alternative column names accepted:**
- `MFAStatus` instead of `MFAEnabled`

### Sample Data

```csv
UserPrincipalName,DisplayName,MFAEnabled
john.doe@company.com,John Doe,True
jane.smith@company.com,Jane Smith,True
bob.jones@company.com,Bob Jones,False
```

---

## User Security Groups CSV Format

This file contains user group memberships, particularly for identifying admin roles.

### Required Columns

| Column Name | Description | Example |
|-------------|-------------|---------|
| UserPrincipalName | User's email address | john.doe@company.com |
| DisplayName | User's display name | John Doe |
| GroupName | Security group name | Domain Admins |

### Sample Data

```csv
UserPrincipalName,DisplayName,GroupName
john.doe@company.com,John Doe,Domain Admins
john.doe@company.com,John Doe,IT Department
jane.smith@company.com,Jane Smith,Sales Team
bob.jones@company.com,Bob Jones,Exchange Administrators
```

**Note:** Users may appear multiple times if they belong to multiple groups.

---

## Interactive SignIns CSV Format

This file contains historical sign-in data for behavioral analysis.

### Required Columns

| Column Name | Description | Example |
|-------------|-------------|---------|
| UserPrincipalName | User's email address | john.doe@company.com |
| DisplayName | User's display name | John Doe |
| IPAddress | Sign-in IP address | 192.168.1.100 |
| Location | Geographic location | New York, NY, USA |
| SignInDateTime | Date and time of sign-in (optional) | 2025-01-15 14:30:00 |

### Sample Data

```csv
UserPrincipalName,DisplayName,IPAddress,Location,SignInDateTime
john.doe@company.com,John Doe,192.168.1.100,New York NY USA,2025-01-15 14:30:00
john.doe@company.com,John Doe,192.168.1.101,New York NY USA,2025-01-15 09:15:00
john.doe@company.com,John Doe,10.0.0.50,London UK,2025-01-14 16:45:00
jane.smith@company.com,Jane Smith,172.16.0.10,London UK,2025-01-15 10:00:00
jane.smith@company.com,Jane Smith,172.16.0.11,London UK,2025-01-14 11:20:00
```

**Note:** The more historical sign-in data you provide, the better the behavioral analysis will be.

---

## How to Export These Files

### From Microsoft 365 / Azure AD

#### MFA Status Export
```powershell
# Using Microsoft Graph PowerShell
Connect-MgGraph -Scopes "UserAuthenticationMethod.Read.All"
$users = Get-MgUser -All
$mfaData = foreach ($user in $users) {
    $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id
    [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        DisplayName = $user.DisplayName
        MFAEnabled = ($authMethods.Count -gt 1).ToString()
    }
}
$mfaData | Export-Csv -Path "MFAStatus.csv" -NoTypeInformation
```

#### User Security Groups Export
```powershell
# Using Microsoft Graph PowerShell
Connect-MgGraph -Scopes "Directory.Read.All"
$users = Get-MgUser -All
$groupData = foreach ($user in $users) {
    $groups = Get-MgUserMemberOf -UserId $user.Id
    foreach ($group in $groups) {
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            GroupName = $group.AdditionalProperties.displayName
        }
    }
}
$groupData | Export-Csv -Path "UserSecurityGroups.csv" -NoTypeInformation
```

#### Interactive SignIns Export
```powershell
# Using Microsoft Graph PowerShell
Connect-MgGraph -Scopes "AuditLog.Read.All"
$startDate = (Get-Date).AddDays(-30)
$signIns = Get-MgAuditLogSignIn -Filter "createdDateTime ge $($startDate.ToString('yyyy-MM-dd'))" -All
$signInData = $signIns | Where-Object { $_.IsInteractive -eq $true } | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = $_.UserPrincipalName
        DisplayName = $_.UserDisplayName
        IPAddress = $_.IPAddress
        Location = "$($_.Location.City), $($_.Location.State), $($_.Location.CountryOrRegion)"
        SignInDateTime = $_.CreatedDateTime
    }
}
$signInData | Export-Csv -Path "InteractiveSignIns.csv" -NoTypeInformation
```

### From Barracuda XDR

1. Log into Barracuda XDR Console
2. Navigate to **Alerts** or **Incidents**
3. Apply desired filters (date range, severity, etc.)
4. Click **Export** and choose **Excel (XLSX)** format
5. Save the file and ensure it has the required columns (you may need to rename columns to match the expected format)

**Column Mapping for Barracuda XDR:**
- Ticket/Incident ID → `TicketNumber`
- User/Account → `User`
- Alert Name/Type → `AlertType`
- Source IP → `IPAddress`
- Geographic Location → `Location`
- VPN/ISP Info → `VPNProvider`

---

## Tips for Best Results

1. **Use Consistent Date Ranges**: Ensure your SignIns CSV covers a reasonable period (30-90 days recommended) for accurate behavioral analysis.

2. **Clean Your Data**: Remove any test accounts or service accounts from the CSV files to avoid skewing the analysis.

3. **Column Headers**: The script is case-insensitive for column headers, but spelling must match exactly.

4. **Optional Fields**: If you don't have some CSV files, that's okay! The script will still generate reports but with less detailed analysis.

5. **Update Regularly**: Keep your MFA status and Security Groups CSVs up to date for accurate risk assessment.

---

## Troubleshooting Data Issues

### "Column not found" errors
- **Cause**: Column names don't match expected format
- **Solution**: Verify column headers match exactly (case-insensitive but spelling must match)

### "No matching users found"
- **Cause**: UserPrincipalName values don't match between files
- **Solution**: Ensure consistent email formats across all files (e.g., all lowercase)

### Incorrect classifications
- **Cause**: Insufficient or inaccurate historical data
- **Solution**: Provide more comprehensive SignIns data (30+ days recommended)

---

## Example Complete Dataset

Here's a minimal complete example for testing:

**VScanAlerts.xlsx** (Sheet1):
```
TicketNumber    User                    AlertType           IPAddress       Location            VPNProvider
INC001          alice@company.com       Impossible Travel   1.2.3.4         Paris, France       NordVPN
```

**MFAStatus.csv**:
```csv
UserPrincipalName,DisplayName,MFAEnabled
alice@company.com,Alice Johnson,True
```

**UserSecurityGroups.csv**:
```csv
UserPrincipalName,DisplayName,GroupName
alice@company.com,Alice Johnson,Domain Users
```

**InteractiveSignIns.csv**:
```csv
UserPrincipalName,DisplayName,IPAddress,Location
alice@company.com,Alice Johnson,1.2.3.4,Paris France
alice@company.com,Alice Johnson,1.2.3.5,Paris France
alice@company.com,Alice Johnson,5.6.7.8,London UK
alice@company.com,Alice Johnson,9.10.11.12,New York USA
alice@company.com,Alice Johnson,1.2.3.4,Paris France
alice@company.com,Alice Johnson,5.6.7.8,London UK
```

This dataset would result in Alice's alert being classified as "Authorized Activity" because she has multiple sign-ins from diverse locations (VPN user pattern).
