# Archive

Scripts archived for reference. Not part of the main VScanMagic workflow.

## Test-EmailWithOneDriveLinks.ps1

Generates a vulnerability scan follow-up email template with placeholders (or auto-created links via Microsoft Graph) for OneDrive share links.

**Note:** Auto-creation of OneDrive links via Microsoft Graph requires tenant admin approval (Files.ReadWrite scope), which not all users will have. The script falls back to manual placeholders when Graph is unavailable.

## Get-OneDriveShareLink.ps1

Helper script used by Test-EmailWithOneDriveLinks.ps1 to create OneDrive share links via Microsoft Graph API.

**Usage from archive:**
```powershell
cd c:\git\vscanmagic\archive
.\Test-EmailWithOneDriveLinks.ps1 -FolderPath "K:\OneDrive\...\2026 - Q1"
```

## company-review-tests/

Company Review API test scripts archived during single-item array fix work (2026-03). Run from project root.

- **Test-CompanyReviewAPI.ps1**, **Test-CompanyReviewSection3.ps1**, **Test-CompanyReviewServerSide.ps1** — Company Review API tests
- **Test-DiscoverySettingsAPI.ps1**, **Test-ExternalAssetsAPI.ps1**, **Test-ExternalAssetsByCompany.ps1**, **Test-ExternalAssetsClientFilter.ps1**, **Test-JobsViewAPI.ps1** — Discovery settings and external assets tests

## Portal capture scripts

Browser console scripts and guides for capturing ConnectSecure portal API calls (report creation, company review, firewall). Use when debugging API differences vs portal behavior.

- **Capture-PortalCompanyReview.js**, **Capture-PortalCompanyReview-DEBUG.js** — Company Review capture
- **Capture-PortalReportCreation.js** — Report creation capture
- **Capture-PortalFirewall.js** — Firewall API capture (device count, vendor types)
- **Capture-PortalWebData.js**, **Capture-PortalProbeNmap.js**, **Capture-PortalOffline.js** — Web data, probe nmap, offline capture
- **CAPTURE-PORTAL-*.md** — Usage guides (COMPANY-REVIEW, REPORT, FIREWALL, WEB-DATA, PROBE-NMAP, OFFLINE)
