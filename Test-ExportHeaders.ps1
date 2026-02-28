# Quick test that Export uses API headers
. .\ConnectSecure-API.ps1
. .\Generate-Report-Example.ps1
Initialize-ConnectSecureApi -BaseUrl (Get-ConnectSecureBaseUrl) -ClientId (Get-ConnectSecureClientId) -ClientSecret (Get-ConnectSecureClientSecret) -ErrorAction SilentlyContinue
$v = Get-ConnectSecureVulnerabilities -Limit 2 -Raw
if ($v -and $v.Count -gt 0) {
    Export-ConnectSecureDataToExcel -Data $v -OutputPath 'test_api_headers.xlsx' -SheetName 'Test'
    Write-Host 'Exported OK'
} else {
    Write-Host 'No data or API not configured - check credentials'
}
