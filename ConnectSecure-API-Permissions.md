# ConnectSecure API Permissions Required

Based on the API endpoints used by VScanMagic, here are the permissions you need in ConnectSecure:

## Required API Permissions

### 1. **Authentication Permission**
- **Endpoint**: `/w/authorize`
- **Permission**: Ability to authenticate and obtain access tokens
- **Note**: This is typically granted automatically when you create an API key

### 2. **Company Data Read Permission**
- **Endpoint**: `/r/company/companies`
- **Permission**: Read access to company/tenant data
- **Used for**: Listing companies to select which one to generate reports for

### 3. **Vulnerability Data Read Permissions**
The following endpoints require read access to vulnerability/report data:

#### a. **Application Vulnerabilities**
- **Endpoint**: `/r/report_queries/application_vulnerabilities`
- **Permission**: Read access to application vulnerability data
- **Used for**: Generating "All Vulnerabilities Report" and "Pending EPSS Report"

#### b. **External Asset Vulnerabilities**
- **Endpoint**: `/r/report_queries/external_asset_vulnerabilities`
- **Permission**: Read access to external vulnerability data
- **Used for**: Generating "External Vulnerabilities Report"

#### c. **Suppressed Vulnerabilities**
- **Endpoint**: `/r/report_queries/application_vulnerabilities_suppressed`
- **Permission**: Read access to suppressed vulnerability data
- **Used for**: Generating "Suppressed Vulnerabilities Report"

#### d. **Remediation Plan**
- **Endpoint**: `/r/asset/get_asset_remediation_plan`
- **Permission**: Read access to remediation plan data
- **Used for**: Generating "Executive Summary Report"

## How to Check/Set Permissions in ConnectSecure

1. **Log into ConnectSecure Portal**
   - Go to: `https://pod0.myconnectsecure.com` (or your pod URL)

2. **Navigate to API Key Settings**
   - Go to: **Global > Settings > Users**
   - Click the **three-dot menu (Action)** next to your user account
   - Select **API Key**

3. **Check API Key Permissions**
   - Look for a "Permissions" or "Scopes" section
   - Ensure the following permissions/scopes are enabled:
     - `read:companies` or `company:read`
     - `read:vulnerabilities` or `vulnerability:read`
     - `read:reports` or `report:read`
     - `read:assets` or `asset:read`

4. **If Permissions Are Not Visible**
   - Some ConnectSecure instances may not show granular permissions
   - The API key might inherit permissions from your user account
   - Check your user account's role/permissions:
     - Go to: **Global > Settings > Users > [Your User]**
     - Check your assigned role (e.g., Administrator, Security Analyst, etc.)
     - Ensure your role has access to:
       - Company data
       - Vulnerability reports
       - Asset data

## Common Permission Issues

### Issue: "Failed to create customer" Error
- **Possible Cause**: API key or user account lacks necessary permissions
- **Solution**: 
  1. Verify your user account has appropriate role/permissions
  2. Check if API key needs to be regenerated with correct permissions
  3. Contact ConnectSecure support if permissions seem correct but error persists

### Issue: "Unauthorized" or "403 Forbidden" Errors
- **Possible Cause**: Missing read permissions for specific endpoints
- **Solution**: 
  1. Verify read permissions for companies, vulnerabilities, and assets
  2. Check if your user role has access to the data you're trying to retrieve
  3. Ensure you're using the correct company ID (if required)

### Issue: Empty Results or No Data Returned
- **Possible Cause**: Permissions exist but are scoped to specific companies/data
- **Solution**: 
  1. Verify you have access to the company ID you're querying
  2. Check if there are any data filters or restrictions on your account
  3. Try with `companyId=0` to see all accessible companies

## Recommended User Role

For best results, your ConnectSecure user account should have:
- **Role**: Administrator or Security Analyst
- **Permissions**: 
  - Full read access to vulnerability data
  - Read access to company/tenant data
  - Read access to asset and remediation data

## Contact ConnectSecure Support

If you're unsure about permissions or continue to experience issues:
1. Contact ConnectSecure support
2. Provide them with:
   - Your tenant name (email address)
   - The error message you're receiving
   - The API endpoints you're trying to access
   - Your user role/permissions

## API Endpoints Summary

| Endpoint | Method | Purpose | Required Permission |
|----------|--------|---------|-------------------|
| `/w/authorize` | POST | Authenticate | Authentication (automatic) |
| `/r/company/companies` | GET | List companies | `read:companies` |
| `/r/report_queries/application_vulnerabilities` | GET | Get vulnerabilities | `read:vulnerabilities` |
| `/r/report_queries/external_asset_vulnerabilities` | GET | Get external vulnerabilities | `read:vulnerabilities` |
| `/r/report_queries/application_vulnerabilities_suppressed` | GET | Get suppressed vulnerabilities | `read:vulnerabilities` |
| `/r/asset/get_asset_remediation_plan` | GET | Get remediation plan | `read:assets` |
