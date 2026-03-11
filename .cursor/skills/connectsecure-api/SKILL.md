---
name: connectsecure-api
description: Integrates with the ConnectSecure vulnerability management API. Use when implementing or debugging ConnectSecure API calls, authentication, report downloads, rate limiting, or when working with report_queries endpoints.
---

# ConnectSecure API Integration

## Authentication

**Endpoint:** `POST /w/authorize`

**Auth format:** Base64 of `tenantname+ClientId:ClientSecret` (UTF-8 encoding). Use `Client-Auth-Token` header.

```powershell
$authString = "${TenantName}+${ClientId}:${ClientSecret}"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$base64Auth = [System.Convert]::ToBase64String($bytes)
# Header: "Client-Auth-Token" = $base64Auth
```

**Required credentials:** BaseUrl, TenantName, ClientId, ClientSecret. Trim whitespace; strip accidental newlines from pasted secrets.

**Response:** Returns `access_token`, `user_id`. Store both. Use `Authorization: Bearer <token>` and `X-USER-ID: <user_id>` on subsequent requests.

**401 handling:** Token may expire. Re-authenticate and retry. Check response body for API-provided error details.

## Rate Limiting

Default limits (per `ConnectSecure-API.ps1`):
- 300 requests/minute
- 2000 requests/hour
- 30000 requests/day

**Flow:** Before each request: `Wait-ForRateLimit` → `Add-RequestToHistory` → make request.

**429 response:** Wait 60 seconds, retry.

## Report Endpoints

| Report | Endpoint | Notes |
|--------|----------|-------|
| All Vulnerabilities | `/r/report_queries/asset_wise_vulnerabilities` | One row per host+vuln; high volume. Use `company_id` to reduce. |
| vulnerabilities_details | `/r/report_queries/vulnerabilities_details` | One row per unique vuln; compact. Used for EPSS, Top 10. |
| Suppressed | `/r/report_queries/vulnerabilities_details_suppressed` | company_id |
| External Scan | `/r/report_queries/external_asset_vulnerabilities` | company_id |
| Registry | `/r/report_queries/registry_problems_remediation` | Uses `condition` param |
| Network | `/r/report_queries/application_vulnerabilities_net` | Uses `condition` (e.g. `software_type='networksoftware'`) |

## Common Query Parameters

| Param | Purpose |
|-------|---------|
| `company_id` | Filter by company. 0 = all. **Use specific company to reduce rows** (e.g. 110k → 10–20k). |
| `limit` | Page size |
| `skip` | Offset for pagination |
| `sort` | e.g. `severity.keyword:desc` |
| `filter` | Optional, e.g. `severity.keyword:(Critical OR High)` |
| `condition` | For registry/network endpoints (not company_id). Example: `company_id=X and is_suppressed=false` |
| `order_by` | For some endpoints, e.g. `affected_assets desc` |

## Path Prefixes

- `/r/` – read
- `/w/` – write/create
- `/d/` – delete

## Helper Functions (ConnectSecure-API.ps1)

| Function | Purpose |
|----------|---------|
| `Connect-ConnectSecureAPI` | Authenticate; stores token in `$script:ConnectSecureConfig` |
| `Invoke-ConnectSecureRequest` | Wraps requests with rate limit, retry, token refresh |
| `Load-ConnectSecureCredentials` | Load saved creds from VScanMagic settings |
| `Write-CSApiLog` | Log with level (Info, Warning, Error, Success) |
| `Remove-SensitiveDataFromObject` | Redact tokens/secrets before logging |

## Standard Reports (Report Builder)

For pre-built report downloads: `GET /report_builder/standard_reports?isGlobal=false`. Report IDs vary by tenant. Data APIs above generate reports programmatically (preferred in VScanMagic).

## Additional Reference

- `ConnectSecure-Vulnerability-Endpoints.md` – Full endpoint list
- `ConnectSecure-Standard-Reports-Reference.md` – Report IDs, data APIs
- `ConnectSecure-API-401-Troubleshooting.md` – Auth debugging
