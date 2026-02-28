$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$lineNum = 1
$col = 0
foreach ($c in $content.ToCharArray()) {
    $code = [int]$c
    if ($code -in 0x22, 0x27, 0x2018, 0x2019, 0x201C, 0x201D, 0x91, 0x92, 0x93, 0x94, 0xFF02) {
        $name = switch ($code) {
            0x22 { 'ASCII-DQ' }
            0x27 { 'ASCII-SQ' }
            0x2018 { 'U+2018' }
            0x2019 { 'U+2019' }
            0x201C { 'U+201C' }
            0x201D { 'U+201D' }
            0x91 { 'CP1252-91' }
            0x92 { 'CP1252-92' }
            0x93 { 'CP1252-93' }
            0x94 { 'CP1252-94' }
            0xFF02 { 'U+FF02-Fullwidth' }
        }
        Write-Host "Line $lineNum Col $col : $name (U+$($code.ToString('X4')))"
    }
    if ($c -eq "`n") { $lineNum++; $col = 0 } else { $col++ }
}
