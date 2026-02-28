$script = Get-Content "c:\Git\vscanmagicv3\ConnectSecure-API.ps1" -Raw
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($script, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    $parseErrors | Select-Object -First 3 | ForEach-Object {
        Write-Host "Error: $($_.Message)"
        Write-Host "  At line $($_.Extent.StartLineNumber) col $($_.Extent.StartColumn)"
        Write-Host "  Text: $($_.Extent.Text)"
    }
} else {
    Write-Host "No parse errors"
}
