# Fix potential UTF-8 encoded smart quotes in raw bytes
$path = "c:\Git\vscanmagicv3\ConnectSecure-API.ps1"
$bytes = [System.IO.File]::ReadAllBytes($path)
$newBytes = [System.Collections.Generic.List[byte]]::new()
$i = 0
$replacements = 0
while ($i -lt $bytes.Count) {
    # UTF-8 for U+201C (") is E2 80 9C
    if ($i -le $bytes.Count - 3 -and $bytes[$i] -eq 0xE2 -and $bytes[$i+1] -eq 0x80 -and $bytes[$i+2] -eq 0x9C) {
        $newBytes.Add(0x22)
        $i += 3
        $replacements++
    }
    # UTF-8 for U+201D (") is E2 80 9D
    elseif ($i -le $bytes.Count - 3 -and $bytes[$i] -eq 0xE2 -and $bytes[$i+1] -eq 0x80 -and $bytes[$i+2] -eq 0x9D) {
        $newBytes.Add(0x22)
        $i += 3
        $replacements++
    }
    # UTF-8 for U+2018 (') is E2 80 98
    elseif ($i -le $bytes.Count - 3 -and $bytes[$i] -eq 0xE2 -and $bytes[$i+1] -eq 0x80 -and $bytes[$i+2] -eq 0x98) {
        $newBytes.Add(0x27)
        $i += 3
        $replacements++
    }
    # UTF-8 for U+2019 (') is E2 80 99
    elseif ($i -le $bytes.Count - 3 -and $bytes[$i] -eq 0xE2 -and $bytes[$i+1] -eq 0x80 -and $bytes[$i+2] -eq 0x99) {
        $newBytes.Add(0x27)
        $i += 3
        $replacements++
    }
    else {
        $newBytes.Add($bytes[$i])
        $i++
    }
}
[System.IO.File]::WriteAllBytes($path, [byte[]]$newBytes)
Write-Host "Fixed $replacements quote characters (raw byte replacement)"
