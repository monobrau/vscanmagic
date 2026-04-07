# VScanMagic API Server

REST API server for generating and downloading vulnerability reports from VScanMagic.

## Features

The API provides endpoints to generate and download the following reports:

1. **Pending EPSS Report** (XLSX) - Excel report with EPSS scores and pivot tables
2. **Executive Summary** (DOCX) - Word document with executive-level vulnerability summary
3. **All Vulnerabilities Report** (XLSX) - Complete list of all vulnerabilities
4. **External Vulnerabilities Report** (XLSX) - Vulnerabilities exposed to external networks
5. **Suppressed Vulnerabilities Report** (XLSX) - List of suppressed vulnerabilities

### Data Sources

Reports can be generated from two data sources:

- **Excel File Upload** - Upload an Excel file containing vulnerability scan data
- **ConnectSecure API** - Fetch data directly from ConnectSecure API (requires API credentials)

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Excel installed
- Microsoft Word installed (for Executive Summary)
- `VScanMagic-ApiBootstrap.ps1` in the same directory (loads Core, Data, Reports). If it is missing, the server falls back to dot-sourcing `VScanMagic-GUI.ps1` (full GUI stack).

## Starting the Server

```powershell
.\VScanMagic-API.ps1
```

The server binds to **loopback** by default: `http://127.0.0.1:8080/` (not reachable from other machines).

### Security (recommended)

| Setting | Purpose |
|--------|---------|
| `VSCANMAGIC_API_KEY` | If set, every request must send the same value in header `X-VScanMagic-Api-Key` or `Authorization: Bearer <key>`. OPTIONS preflight is exempt. |
| `VSCANMAGIC_API_BIND` | Optional. Override bind host (default `127.0.0.1`). Do not expose this API to untrusted networks without authentication. |

- **Do not pass** `ConnectSecureClientSecret` in the query string; **JSON body only** (query strings appear in logs and browser history).
- CORS allows `*` for local development; treat as **trusted local use only**.

## API Endpoints

### Health Check

**GET** `/health`

Returns server status and version information.

**Response:**
```json
{
  "status": "healthy",
  "version": "1.0.0",
  "timestamp": "2026-02-23 10:30:00"
}
```

### Generate Reports

**POST** `/api/reports/{report-type}`

Generates a report and returns it as a file download.

**Report Types:**
- `pending-epss` - Pending EPSS Report (XLSX)
- `executive-summary` - Executive Summary (DOCX)
- `all-vulnerabilities` - All Vulnerabilities Report (XLSX)
- `external-vulnerabilities` - External Vulnerabilities Report (XLSX)
- `suppressed-vulnerabilities` - Suppressed Vulnerabilities Report (XLSX)

**Request Body (JSON) - Excel File Mode:**
```json
{
  "inputPath": "C:\\Path\\To\\Input\\File.xlsx",
  "clientName": "Client Name",
  "scanDate": "02/23/2026"
}
```

**Request Body (JSON) - ConnectSecure API Mode:**
```json
{
  "useConnectSecure": true,
  "connectSecureBaseUrl": "https://pod0.myconnectsecure.com",
  "connectSecureTenant": "your-tenant-name",
  "connectSecureClientId": "your-client-id",
  "connectSecureClientSecret": "your-client-secret",
  "companyId": 123,
  "clientName": "Client Name",
  "scanDate": "02/23/2026"
}
```

**Query Parameters (alternative to JSON body):**
- `inputPath` (required for Excel mode) - Path to the input Excel file
- `useConnectSecure` (optional) - Set to `true` to use ConnectSecure API
- `connectSecureBaseUrl` (required for ConnectSecure mode) - ConnectSecure API base URL (e.g., `https://pod0.myconnectsecure.com`)
- `connectSecureTenant` (required for ConnectSecure mode) - Your ConnectSecure tenant name
- `connectSecureClientId` (required for ConnectSecure mode) - Your ConnectSecure API client ID
- `connectSecureClientSecret` (required for ConnectSecure mode) - Your ConnectSecure API client secret
- `companyId` (optional) - ConnectSecure company ID (0 for all companies)
- `clientName` (optional) - Client name for report (default: "Client")
- `scanDate` (optional) - Scan date in MM/dd/yyyy format (default: current date)

**Response:**
- Success: File download (Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet or .wordprocessingml.document)
- Error: JSON error message with HTTP status code

**Error Response Example:**
```json
{
  "error": "InputPath parameter is required"
}
```

## Usage Examples

### Using PowerShell Invoke-RestMethod

**Excel File Mode:**
```powershell
# Generate Pending EPSS Report from Excel file
$body = @{
    inputPath = "C:\Reports\VulnerabilityScan.xlsx"
    clientName = "Acme Corp"
    scanDate = "02/23/2026"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:8080/api/reports/pending-epss" `
    -Method POST `
    -Body $body `
    -ContentType "application/json" `
    -OutFile "PendingEPSSReport.xlsx"
```

**ConnectSecure API Mode:**
```powershell
# Generate Pending EPSS Report from ConnectSecure API
$body = @{
    useConnectSecure = $true
    connectSecureBaseUrl = "https://pod0.myconnectsecure.com"
    connectSecureTenant = "your-tenant-name"
    connectSecureClientId = "your-client-id"
    connectSecureClientSecret = "your-client-secret"
    companyId = 123
    clientName = "Acme Corp"
    scanDate = "02/23/2026"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:8080/api/reports/pending-epss" `
    -Method POST `
    -Body $body `
    -ContentType "application/json" `
    -OutFile "PendingEPSSReport.xlsx"
```

### Using cURL

```bash
# Generate Executive Summary
curl -X POST "http://localhost:8080/api/reports/executive-summary" \
  -H "Content-Type: application/json" \
  -d '{"inputPath":"C:\\Reports\\VulnerabilityScan.xlsx","clientName":"Acme Corp","scanDate":"02/23/2026"}' \
  --output ExecutiveSummary.docx
```

### Using Python

```python
import requests

url = "http://localhost:8080/api/reports/all-vulnerabilities"
data = {
    "inputPath": r"C:\Reports\VulnerabilityScan.xlsx",
    "clientName": "Acme Corp",
    "scanDate": "02/23/2026"
}

response = requests.post(url, json=data)
if response.status_code == 200:
    with open("AllVulnerabilities.xlsx", "wb") as f:
        f.write(response.content)
    print("Report downloaded successfully")
else:
    print(f"Error: {response.json()}")
```

### Using JavaScript/Node.js

```javascript
const axios = require('axios');
const fs = require('fs');

async function generateReport() {
    try {
        const response = await axios.post(
            'http://localhost:8080/api/reports/external-vulnerabilities',
            {
                inputPath: 'C:\\Reports\\VulnerabilityScan.xlsx',
                clientName: 'Acme Corp',
                scanDate: '02/23/2026'
            },
            {
                responseType: 'arraybuffer'
            }
        );
        
        fs.writeFileSync('ExternalVulnerabilities.xlsx', response.data);
        console.log('Report downloaded successfully');
    } catch (error) {
        console.error('Error:', error.response?.data || error.message);
    }
}

generateReport();
```

## Configuration

Edit the `$script:ApiConfig` hashtable in `VScanMagic-API.ps1` to change:

- `Port` - Server port (default: 8080)
- `Host` - Server hostname (default: "localhost")
- `TempDirectory` - Temporary directory for generated files (default: `$env:TEMP\VScanMagic-API`)

## CORS Support

The API includes CORS headers to allow cross-origin requests from web applications.

## ConnectSecure API Integration

The API supports fetching vulnerability data directly from ConnectSecure API v4. This eliminates the need to manually export Excel files.

### Getting ConnectSecure API Credentials

1. Log in to your ConnectSecure portal
2. Navigate to **Global > Settings > Users**
3. Click the three-dot menu (Action) next to your user
4. Select **API Key**
5. Copy your **Client ID** and **Client Secret**
6. Your **Tenant Name** is your ConnectSecure tenant identifier
7. Your **Base URL** is typically `https://pod{number}.myconnectsecure.com` (check your portal URL)

### Rate Limits

The ConnectSecure API has the following rate limits:
- 300 requests per minute
- 2,000 requests per hour
- 30,000 requests per day

The API client automatically handles rate limiting and will wait when limits are reached.

### Authentication

The API uses base64-encoded authentication tokens in the format:
```
base64(tenant+client_id:client_secret)
```

The authentication is handled automatically by the API client.

## Notes

- Generated files are temporarily stored in the configured temp directory
- Files are automatically cleaned up 60 seconds after download
- The server must have access to the input Excel file path (for Excel mode)
- Excel and Word COM objects are properly cleaned up after each request
- The server runs synchronously - each request is processed sequentially
- ConnectSecure API requests are rate-limited automatically
- ConnectSecure authentication tokens are cached and refreshed automatically

## Troubleshooting

**Error: "VScanMagic-GUI.ps1 not found"**
- Ensure `VScanMagic-GUI.ps1` is in the same directory as `VScanMagic-API.ps1`

**Error: "Failed to create Excel COM object"**
- Ensure Microsoft Excel is installed and properly registered

**Error: "Input file not found"**
- Verify the `inputPath` is correct and accessible
- Use absolute paths, not relative paths
- Ensure the file is not locked by another process

**Port already in use**
- Change the port in `$script:ApiConfig.Port`
- Or stop the process using port 8080
