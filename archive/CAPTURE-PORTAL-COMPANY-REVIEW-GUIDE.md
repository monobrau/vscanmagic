# Capturing ConnectSecure Portal Company Review API Calls

Capture the exact API calls the ConnectSecure web UI makes when viewing company data (agents, assets, external scans, etc.). Use this to discover **server-side filtering** parameters (e.g. `company_id`, `condition`) so we can filter data on the server instead of client-side.

---

**IMPORTANT: Always capture from the web portal before implementing client-side filtering.**  
When adding or changing API calls that filter by company (or other criteria), first use this guide to capture what the ConnectSecure portal does. The portal may use server-side params (`condition`, `company_id`, etc.) that work—adopting those avoids slow "fetch all, filter locally" patterns. Only fall back to client-side filtering when the portal capture shows no working server-side option.

## Steps

1. **Log into ConnectSecure** in Chrome or Edge
2. **Open DevTools** (F12) → go to the **Console** tab
3. **Paste the contents** of `Capture-PortalCompanyReview.js` into the console and press Enter
4. You should see: `[Company Review Capturer Active] Navigate to a company view...`
5. **Navigate to company data** in the portal:
   - Select a **specific company** (not Global) from the company selector
   - Go to pages that show:
     - **Agents** (or Agent list)
     - **Assets** (or Asset list)
     - **External Scans** / **External Assets** (if available)
     - **Discovery Settings** / **Credentials** (if visible)
   - Let each page load fully so the portal makes its API calls
6. **Export captured data**:
   - Run in console: `copy(JSON.stringify(window.__capturedCompanyCalls, null, 2))`
   - This copies the captured calls to your clipboard
   - Paste into a file: `portal-company-review-capture.json`

## What to Look For

For each captured request, check:

- **`queryParams`** – Does the portal send `company_id`, `condition`, or other filters?
- **`path`** – Exact endpoint path (e.g. `/r/report_queries/external_asset_externalscan`)
- **`responseRecordCount`** – Number of records returned (helps verify filtering)
- **`requestBody`** – For POST requests, what body does the portal send?

### Key Endpoints

| Endpoint | Purpose | Current VScanMagic |
|----------|---------|--------------------|
| `/r/report_queries/discovery_settings` | External scan config | Server-side `condition=company_id=X` |
| `/r/report_queries/lightweight_assets` | Lightweight agents | `company_id` query param |
| `/r/company/agents` | All agents | Server-side `condition=company_id=X` |
| `/r/company/jobs_view` | Scan jobs (last external scan date) | Server-side `condition=company_id=X and type='External Scan'` |
| `/r/company/credentials`, `agent_credentials_mapping`, `agent_discoverysettings_mapping`, `asset_firewall_policy`, `company_stats` | Various | Server-side `condition=company_id=X` |

If the portal uses `company_id=X` or `condition=company_ids:X` successfully, we can adopt that for server-side filtering.

## Alternative: DevTools Network Tab

1. **Open DevTools** (F12) → **Network** tab
2. Enable **Preserve log**
3. Filter by **Fetch/XHR**
4. Select a company and navigate to Agents, Assets, External Scans
5. **Inspect requests** – Click each request to see:
   - **Headers** → Request URL (includes query string)
   - **Payload** (for POST)
   - **Response** → record count

## Portal Capture Findings

### discovery_settings (2026-03-01)

When viewing external scan config or External Assets, the portal calls:

```
GET https://api.myconnectsecure.com/r/report_queries/discovery_settings
  condition: "company_id=15373 and discovery_settings_type='externalscan'"
  skip: 0, limit: 100, order_by: "updated desc"
```

- **Endpoint**: `/r/report_queries/discovery_settings` (not `/r/company/discovery_settings`)
- **Condition format**: `company_id=X` (equals sign), with optional `and discovery_settings_type='externalscan'`
- **Implemented**: `Get-ConnectSecureDiscoverySettings` now tries this endpoint first with `condition=company_id=X`; falls back to `/r/company/discovery_settings` + client filter if needed

## Capturing the Last External Scan Date Source

**Goal:** Find out where the portal gets the "last external scan" date—and whether it requires an assets fetch or uses a simpler endpoint.

1. **Find where the portal shows it** – Look for "Last scanned", "Last scan", "Updated", or similar on:
   - Company dashboard / overview
   - External Assets / External Scans view
   - Any company-level summary
2. **Run the DEBUG capturer** – Paste `Capture-PortalCompanyReview-DEBUG.js` into the console
3. **Refresh the page** (or navigate to the company view) so data loads with the script active
4. **Export all calls** – Run: `copy(JSON.stringify(window.__capturedAllApiCalls, null, 2))`
5. **Inspect** – Search the output for:
   - `company_stats` – does it have `external_last_scan_time`?
   - `external_asset_externalscan` – does the portal call it? With what params?
   - Any endpoint with `scan`, `stats`, or `updated` in the path

If the portal gets the date from `company_stats` (or another single-call endpoint) without fetching assets, we can adopt that. If it uses `external_asset_externalscan` and we see server-side filtering, we can use that.

**jobs_view (implemented):** The portal calls `/r/company/jobs_view`. With `condition=company_id=X and type='External Scan'` it returns external scan jobs server-side. Use max(updated) for last external scan date—no assets fetch. Implemented in Get-ConnectSecureCompanyReviewData.

## Capturing for a Company Where External Assets Don't Show (e.g. Example Company)

When the portal shows external assets for a company but VScanMagic Company Review does not, capture from that company's view to compare.

1. **Find the company ID** – In the portal, select your target company. The URL may contain the company ID (e.g. `company_id=12345` or `company/12345`). Or run:
   ```powershell
   .\Test-ExternalAssetsByCompany.ps1 -ListCompanies
   ```
   Then search the output for your company name and note its ID.

2. **Clear and start capture** – Paste `Capture-PortalCompanyReview-DEBUG.js` into the console. Clear any previous capture: `window.__capturedAllApiCalls = []`

3. **Navigate to the company** – Select the company from the company selector, then go to **External Assets** or **External Scan Endpoints** (where the portal shows their external assets).

4. **Export** – Run: `copy(JSON.stringify(window.__capturedAllApiCalls, null, 2))` and save to `portal-company-capture.json`

5. **Share** – Paste the captured calls (or the `discovery_settings` and `external_asset` related entries) so we can see:
   - What endpoint the portal uses
   - What `condition` or `company_id` it sends
   - The `responseRecordCount` and `firstRecordSample` for discovery_settings

## Use the Capture

Share `portal-company-review-capture.json` (or paste relevant snippets) to compare with `ConnectSecure-API.ps1`. If the portal uses working server-side params, we can update our API calls to use them instead of fetching all data and filtering client-side.
