# Re-save file with UTF-8 BOM (PowerShell on Windows sometimes parses UTF-8 BOM better)
$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
Write-Host "Re-saved with UTF-8 BOM"
