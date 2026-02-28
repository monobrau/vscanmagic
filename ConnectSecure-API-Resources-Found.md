# ConnectSecure API Resources Found

## GitHub Repositories

### 1. **simon-r-watson/ConnectSecureVulnerabilityManagment**
- **URL**: https://github.com/simon-r-watson/ConnectSecureVulnerabilityManagment
- **Status**: Active (1 star, 1 fork, 20 commits)
- **Description**: Previously known as CyberCNS
- **Language**: PowerShell module
- **OpenAPI Spec**: Includes OpenAPI 3.0.2 specification
- **Note**: The OpenAPI spec shows `/api/{tenant}/authorize` as a **GET** endpoint (not POST `/w/authorize`)

### 2. **redanthrax/CyberCNSAPI** (Archived)
- **URL**: https://github.com/redanthrax/CyberCNSAPI
- **Status**: Archived (read-only since April 16, 2022)
- **Description**: Unofficial CyberCNS API PowerShell module
- **Authentication**: Uses username/password credentials (not API keys)
- **Usage Example**:
  ```powershell
  Import-Module .\CyberCNSApi.psm1 -Force
  $securePw = ConvertTo-SecureString "Password123" -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential("tech@acme.com", $securePw)
  Connect-CyberCNSApi -Url "acme.mycybercns.com" -Credential $credential
  ```

## API Endpoint Differences Found

### OpenAPI Spec Shows:
- **Endpoint**: `/api/{tenant}/authorize`
- **Method**: GET
- **Path Parameter**: `tenant` (tenant name)

### Your Current Implementation Uses:
- **Endpoint**: `/w/authorize`
- **Method**: POST
- **Header**: `Client-Auth-Token` (Base64 encoded `tenant+client_id:client_secret`)

**Note**: These appear to be different API versions or different authentication methods. The `/w/authorize` endpoint you're using matches the Swagger UI at `pod104.myconnectsecure.com/apidocs/`, which suggests it's the correct v4 API endpoint.

## Documentation Resources

### Official Documentation
1. **ConnectSecure V4 Resources**: https://cybercns.atlassian.net/wiki/spaces/CVB/pages/2180153345/ConnectSecure+V4+Resources
2. **API Documentation**: Available at `https://pod104.myconnectsecure.com/apidocs/` (Swagger UI)
3. **Support Portal**: https://cybercns.freshdesk.com/

### Integration Guides
- **vCIOToolbox Integration**: https://vciotoolbox.freshdesk.com/support/solutions/articles/43000743443-connectsecure-integration
- **Agent Deployment**: https://cybercns.freshdesk.com/support/solutions/articles/66000531735-how-to-install-v4-agent-using-rmm-script

## Key Findings

1. **Different Authentication Methods**:
   - Older API (CyberCNSAPI repo): Username/password credentials
   - Current API (v4): Client ID/Secret with Base64 token in header

2. **Endpoint Variations**:
   - OpenAPI spec shows: `/api/{tenant}/authorize` (GET)
   - Swagger UI shows: `/w/authorize` (POST)
   - Your implementation uses: `/w/authorize` (POST) ✓

3. **No Public Examples Found**:
   - No public GitHub repositories with working `/w/authorize` POST examples
   - The archived repository uses a different authentication method
   - Most examples found are for agent deployment, not API authentication

## Conclusion

Your implementation appears to be correct based on:
- The Swagger UI at `pod104.myconnectsecure.com/apidocs/` showing `/w/authorize` POST endpoint
- The official documentation mentioning Client ID/Secret authentication
- The "Failed to create customer" error (not "Failed to authorize") confirming credentials are correct

The "Failed to create customer" error is likely a ConnectSecure API-side issue that needs to be resolved with their support team.
