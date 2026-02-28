# Quick Start: Using Your ConnectSecure API Key

## 1. Get Your API Credentials

1. Log in to ConnectSecure portal
2. Go to **Global > Settings > Users**
3. Click **three-dot menu** → **API Key**
4. Copy:
   - **Client ID**
   - **Client Secret**
   - **Tenant Name**
   - **Base URL** (format: `https://pod{number}.myconnectsecure.com`, e.g., `https://pod104.myconnectsecure.com` - POD number from API Key page)
   - **Tenant Name** (e.g., `river-run` - found on API Key page)

## 2. Start the API Server

Open PowerShell and run:

```powershell
.\VScanMagic-API.ps1
```

Keep this terminal open - the server needs to stay running.

## 3. Generate a Report

### Option A: Use the Example Script (Easiest)

1. Open `Generate-Report-Example.ps1` in a text editor
2. Replace the placeholder values in the `$config` hashtable:
   ```powershell
   $config = @{
       BaseUrl = "https://pod104.myconnectsecure.com"  # Format: https://pod{number}.myconnectsecure.com
       TenantName = "river-run"  # Found on API Key page
       TenantName = "your-tenant-name"                  # Your tenant
       ClientId = "your-client-id"                      # From API Key
       ClientSecret = "your-client-secret"              # From API Key
       CompanyId = 0                                     # 0 = all companies
       ClientName = "My Client"
       ApiServerUrl = "http://localhost:8080"
   }
   ```
3. Save the file
4. Run it:
   ```powershell
   .\Generate-Report-Example.ps1
   ```

### Option B: Use PowerShell Directly

```powershell
$body = @{
    useConnectSecure = $true
    connectSecureBaseUrl = "https://pod104.myconnectsecure.com"
    connectSecureTenant = "river-run"
    connectSecureTenant = "your-tenant-name"
    connectSecureClientId = "your-client-id"
    connectSecureClientSecret = "your-client-secret"
    companyId = 0
    clientName = "My Client"
    scanDate = (Get-Date).ToString("MM/dd/yyyy")
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:8080/api/reports/pending-epss" `
    -Method POST -Body $body -ContentType "application/json" `
    -OutFile "Report.xlsx"
```

## 4. Available Reports

Replace `pending-epss` with any of these:

- `pending-epss` → Pending EPSS Report (XLSX)
- `executive-summary` → Executive Summary (DOCX)
- `all-vulnerabilities` → All Vulnerabilities (XLSX)
- `external-vulnerabilities` → External Vulnerabilities (XLSX)
- `suppressed-vulnerabilities` → Suppressed Vulnerabilities (XLSX)

## Troubleshooting

**"API server is not running"**
- Make sure `VScanMagic-API.ps1` is running in another terminal

**"Failed to authenticate"**
- Double-check your credentials (no extra spaces)
- Verify Base URL includes `https://` and correct pod number

**"No vulnerability data found"**
- Try `companyId = 0` to get all companies
- Check ConnectSecure portal to confirm data exists

## Security Note

⚠️ **Never commit your API credentials to Git!**

The example script is already in `.gitignore`, but if you create your own config file, make sure to:
- Add it to `.gitignore`
- Never share it publicly
- Use environment variables in production
