$path = Join-Path $PSScriptRoot 'ConnectSecure-API.ps1'
$content = Get-Content $path -Raw
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    $parseErrors | ForEach-Object {
        "Line $($_.Extent.StartLineNumber) Col $($_.Extent.StartColumnNumber): $($_.Message)"
    }
} else {
    "No parse errors found."
}
