# ConnectSecure Standard Reports Reference

From `GET /report_builder/standard_reports?isGlobal=false`

## VScanMagic Report Types → ConnectSecure Standard Report IDs

VScanMagic generates these reports **locally** from data APIs (`/r/report_queries/*`).  
These IDs are for reference if using the report builder (`create_report_job`) in the future.

| VScanMagic Report | ConnectSecure Title | Report ID (xlsx) | reportType |
|-------------------|---------------------|------------------|------------|
| All Vulnerabilities | All Vulnerabilities Report | `f836d6a4e4d54ac6a9d2967254796373` | xlsx |
| Suppressed Vulnerabilities | Suppressed Vulnerabilities | `1d091564830b44c485a0ddc35ace9ac6` | xlsx |
| External Scan | External Scan | `01beb6b930744e11b690bb9dc25118fb` | xlsx |
| Executive Summary | Executive Summary Report | `1cd4f45884264d15bee4173dc58b6a57` | docx |
| Pending Remediation EPSS | Pending Remediation EPSS Score Reports | `85d4913c0dbc4fc782b858f0d27dd180` | xlsx |

## Data APIs Used (Current Implementation)

Reports are generated from these endpoints - **not** the report builder:

| Report | API Endpoint |
|--------|--------------|
| All Vulnerabilities | `/r/report_queries/vulnerabilities_details` |
| Suppressed Vulnerabilities | `/r/report_queries/vulnerabilities_details_suppressed` |
| External Scan | `/r/report_queries/external_asset_vulnerabilities` |
| Executive Summary | `/r/report_queries/vulnerabilities_details` (Top 10) |
| Pending Remediation EPSS | `/r/report_queries/vulnerabilities_details` (EPSS filter) |

## create_report_job Parameters (Swagger)

If using report builder:
- `reportId` - use the ID from the table above
- `reportType` - xlsx or docx
- `fileType` - xlsx, docx, pdf
- `company_id` - 0 for all companies
- `company_name`, `reportName`, `isFilter`, `reportFilter` - optional

## Related Documentation

- **Vulnerability & Suppress endpoints**: See `ConnectSecure-Vulnerability-Endpoints.md`

## Standard Reports API

- **Endpoint**: `GET /report_builder/standard_reports`
- **Query**: `isGlobal=false` (required), `skip`, `limit`
- **Header**: `X-USER-ID` (required)
