# How to Use Your ConnectSecure API Key

## Step 1: Get Your API Credentials

1. Log in to your ConnectSecure portal
2. Navigate to **Global > Settings > Users**
3. Click the **three-dot menu (Action)** next to your user account
4. Select **API Key**
5. Copy the following:
   - **Client ID**
   - **Client Secret**
   - **Tenant Name** (your tenant identifier, e.g., `river-run` - found on API Key page)
   - **Base URL** (format: `https://pod{number}.myconnectsecure.com`, e.g., `https://pod104.myconnectsecure.com` - POD number found on API Key page)

## Step 2: Use the API

### Option A: Using PowerShell (Recommended)

Create a PowerShell script with your credentials:

```powershell
# Your ConnectSecure API credentials
$connectSecureConfig = @{
    BaseUrl = "https://pod401.myconnectsecure.com"  # Replace with your pod URL
    TenantName = "your-tenant-name"                  # Replace with your tenant
    ClientId = "your-client-id"                      # Replace with your Client ID
    ClientSecret = "your-client-secret"              # Replace with your Client Secret
    CompanyId = 123                                   # Optional: 0 for all companies
}

# API Server URL
$apiUrl = "http://localhost:8080"

# Generate a report
$body = @{
    useConnectSecure = $true
    connectSecureBaseUrl = $connectSecureConfig.BaseUrl
    connectSecureTenant = $connectSecureConfig.TenantName
    connectSecureClientId = $connectSecureConfig.ClientId
    connectSecureClientSecret = $connectSecureConfig.ClientSecret
    companyId = $connectSecureConfig.CompanyId
    clientName = "My Client"
    scanDate = (Get-Date).ToString("MM/dd/yyyy")
} | ConvertTo-Json

# Generate Pending EPSS Report
Invoke-RestMethod -Uri "$apiUrl/api/reports/pending-epss" `
    -Method POST `
    -Body $body `
    -ContentType "application/json" `
    -OutFile "PendingEPSSReport.xlsx"

Write-Host "Report generated successfully!" -ForegroundColor Green
```

### Option B: Using cURL

```bash
curl -X POST "http://localhost:8080/api/reports/pending-epss" \
  -H "Content-Type: application/json" \
  -d '{
    "useConnectSecure": true,
    "connectSecureBaseUrl": "https://pod401.myconnectsecure.com",
    "connectSecureTenant": "your-tenant-name",
    "connectSecureClientId": "your-client-id",
    "connectSecureClientSecret": "your-client-secret",
    "companyId": 123,
    "clientName": "My Client",
    "scanDate": "02/23/2026"
  }' \
  --output PendingEPSSReport.xlsx
```

### Option C: Using Python

```python
import requests
from datetime import datetime

# Your ConnectSecure API credentials
config = {
    "base_url": "https://pod401.myconnectsecure.com",
    "tenant_name": "your-tenant-name",
    "client_id": "your-client-id",
    "client_secret": "your-client-secret",
    "company_id": 123
}

# API request
url = "http://localhost:8080/api/reports/pending-epss"
data = {
    "useConnectSecure": True,
    "connectSecureBaseUrl": config["base_url"],
    "connectSecureTenant": config["tenant_name"],
    "connectSecureClientId": config["client_id"],
    "connectSecureClientSecret": config["client_secret"],
    "companyId": config["company_id"],
    "clientName": "My Client",
    "scanDate": datetime.now().strftime("%m/%d/%Y")
}

response = requests.post(url, json=data)
if response.status_code == 200:
    with open("PendingEPSSReport.xlsx", "wb") as f:
        f.write(response.content)
    print("Report generated successfully!")
else:
    print(f"Error: {response.json()}")
```

## Step 3: Available Report Types

Replace `/api/reports/pending-epss` with any of these endpoints:

- `/api/reports/pending-epss` - Pending EPSS Report (XLSX)
- `/api/reports/executive-summary` - Executive Summary (DOCX)
- `/api/reports/all-vulnerabilities` - All Vulnerabilities Report (XLSX)
- `/api/reports/external-vulnerabilities` - External Vulnerabilities Report (XLSX)
- `/api/reports/suppressed-vulnerabilities` - Suppressed Vulnerabilities Report (XLSX)

## Step 4: Start the API Server

Before making requests, make sure the API server is running:

```powershell
.\VScanMagic-API.ps1
```

The server will start on `http://localhost:8080` by default.

## Security Best Practices

1. **Never commit credentials to version control**
   - Store credentials in environment variables or a secure config file
   - Add `ConnectSecure-Config.ps1` to `.gitignore`

2. **Use environment variables** (recommended):
   ```powershell
   $env:CONNECTSECURE_BASEURL = "https://pod401.myconnectsecure.com"
   $env:CONNECTSECURE_TENANT = "your-tenant-name"
   $env:CONNECTSECURE_CLIENTID = "your-client-id"
   $env:CONNECTSECURE_CLIENTSECRET = "your-client-secret"
   ```

3. **Use a configuration file** (see `Generate-Report-Example.ps1` below)

## Troubleshooting

**Error: "Failed to authenticate with ConnectSecure API"**
- Verify your credentials are correct
- Check that your Base URL includes `https://` and the correct pod number
- Ensure your Client ID and Client Secret match exactly (no extra spaces)

**Error: "Rate limit exceeded (429)"**
- The API automatically handles rate limiting
- Wait a minute and try again
- Consider reducing the number of concurrent requests

**Error: "No vulnerability data found"**
- Check your CompanyId (use 0 for all companies)
- Verify you have vulnerabilities in ConnectSecure for that company
- Check the ConnectSecure portal to confirm data exists
