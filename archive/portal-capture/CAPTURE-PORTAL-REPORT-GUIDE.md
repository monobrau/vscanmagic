# Capturing ConnectSecure Portal Report Creation

Capture the exact API calls the ConnectSecure web UI makes when creating a standard report. Use this to compare with API-based creation and find differences (e.g. why portal-created reports work with `get_report_link` but API-created ones return 404).

## Method 1: Console Script (Recommended)

1. **Log into ConnectSecure** in Chrome or Edge
2. **Open DevTools** (F12) → go to the **Console** tab
3. **Paste the contents** of `Capture-PortalReportCreation.js` into the console and press Enter
4. You should see: `[Report Capturer Active] Create a report in the UI...`
5. **Create a report** in the portal:
   - Go to Reports (or equivalent menu)
   - Select a standard report type
   - Choose company (e.g. 12345)
   - Click Generate / Create / Download
   - Wait for the report to be ready and downloaded
6. **Export captured data**:
   - Run in console: `copy(JSON.stringify(window.__capturedReportCalls, null, 2))`
   - This copies the captured calls to your clipboard
   - Paste into a file (e.g. `portal-report-capture.json`) for analysis

## Method 2: DevTools Network Tab

1. **Open DevTools** (F12) → **Network** tab
2. Enable **Preserve log**
3. Filter by **Fetch/XHR** (or filter for `report`, `create_report`, `get_report`)
4. Create a report in the UI
5. **Right-click** a request → **Copy** → **Copy as cURL** or **Copy as HAR**
6. Inspect:
   - `create_report_job` (or similar) – request body and response
   - `get_report_link` – parameters and response
   - Any polling or status-check calls

## What to Look For

- **Request payload** for create: `company_id`, `company_name`, `reportId`, `reportType`, `reportFilter`, etc.
- **Headers**: `X-USER-ID`, `Authorization`, any custom headers
- **Polling pattern**: How often does the UI poll for status or the download link?
- **Endpoint paths**: `/report_builder/` vs `/r/report_builder/`
- **Job ID**: Where it appears in responses and how it’s used for the download link

## Portal Findings (2026-02-27 capture)

- **create_report_job**: `reportType: "Standard"`, `isFilter: true`, `reportName` without spaces (e.g. AllVulnerabilitiesReport)
- **create_report_job (Global)**: `company_id: "global"` (string), `company_name: "Global"` - VScanMagic updated to use this format when CompanyId=0
- **get_report_link**: `job_id=["uuid"]` (JSON array), `isGlobal=true` for global reports, **no company_id** - scripts updated to try this first
- **global_jobs_view** (polling): `GET /r/company/global_jobs_view?condition=type='StandardReports'&skip=0&limit=100&order_by=updated desc` - lists global report jobs; each has `job_id`, `status` (completed/pushedtoqueue), `job_details[].r2_path`
- **Large reports**: Returns ZIP with CSV when data is large

## Capturing Global Reports (2026-02-28)

To capture how the portal creates **Global** report types (e.g. "Installed Programs - Global"):

1. **Switch to Global view** - Use the company selector to switch to Global
2. **Open Reports** - Go to Reports / Report Builder
3. **Select a Global report** - e.g. "Installed Programs - Global"
4. **Generate the report** - Click Generate/Download
5. **Export** - Run `copy(JSON.stringify(window.__capturedReportCalls, null, 2))` and save to `portal-global-report-capture.json`

Compare: endpoint path, request body (company_id? isGlobal?), and headers.

## Use the Capture

Compare the portal’s payload and sequence with what `Invoke-CreateReportJob.ps1` and `Invoke-CreateAndDownloadReport.ps1` send. Differences may explain why portal jobs work with `get_report_link` and API jobs do not.
