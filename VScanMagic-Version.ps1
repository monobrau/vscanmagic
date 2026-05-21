# VScanMagic GUI application version — single source of truth.
# Bump only this file when releasing; VScanMagic-GUI.ps1, Core, and BuildExeFinal.ps1 read it.
if (-not $script:VScanMagicVersion) {
    $script:VScanMagicVersion = '4.0.11'
}

function Get-VScanMagicVersion {
    return $script:VScanMagicVersion
}
