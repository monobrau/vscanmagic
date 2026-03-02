# ConnectSecure API: 401 on /r/ Data Endpoints

## Problem

- **Auth works**: `/w/authorize` returns 200 and a valid Bearer token
- **Data endpoints fail**: `/r/company/companies` (and likely other `/r/` endpoints) return **401 Unauthorized**
- **Both PowerShell and Python** fail the same way, so this is not a client-side bug

## What's Been Verified

- Authentication format is correct (tenant+client_id:client_secret, UTF8 base64)
- Token is obtained successfully
- Bearer token is sent in the `Authorization` header
- Same credentials, same URL – auth succeeds, companies request fails

## Likely Cause

The 401 on `/r/` endpoints with a valid auth token points to:

1. **API key / role permissions** – The token may not include access to company/data endpoints
2. **ConnectSecure configuration** – Your pod or API key may need explicit enablement for data APIs
3. **Different auth flow for data APIs** – Some ConnectSecure setups may use a different auth path for `/r/` endpoints

## Recommendations

### 1. Check API Key and User Permissions

- Go to **Global > Settings > Users** in ConnectSecure
- Open your user and choose **Action > API Key**
- Confirm **Company Level Access** is set correctly
- Ensure the user role has read access to company and related data

### 2. Test via Swagger UI

1. Open `https://pod0.myconnectsecure.com/apidocs/` in a browser
2. Authenticate via `/w/authorize` to obtain a token
3. Call the companies endpoint (e.g. `/r/company/companies`)
4. If Swagger returns 401 as well, the issue is on the ConnectSecure side

### 3. Contact ConnectSecure Support

Since auth works but data endpoints return 401:

- **Support**: https://cybercns.freshdesk.com/ or support@connectsecure.com
- **Context**: `/w/authorize` succeeds; `/r/company/companies` returns 401
- **Ask**: Whether your API key has access to company/data APIs and if any extra configuration is needed for `/r/` endpoints

### 4. Confirm Endpoint Path and Base URL

- Confirm your base URL is correct (e.g. `https://pod0.myconnectsecure.com`)
- If you have access to ConnectSecure’s API docs, verify the exact path for the companies endpoint (e.g. `/r/company/companies` vs. another variant)

## Reference

- User Management / API Key: https://cybercns.atlassian.net/wiki/spaces/CVB/pages/2111111353/User+Management+and+Security
- Stellar Cyber connector (401 vs 403): 401 = invalid credentials or insufficient privileges
