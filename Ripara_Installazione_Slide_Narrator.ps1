$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot
$appDataRoot = Join-Path $env:LOCALAPPDATA "SlideNarrator"
Write-Host "Questa procedura elimina soltanto il runtime Python di Slide Narrator." 
Write-Host "Il codice, i PowerPoint e gli altri documenti non vengono cancellati."
Write-Host ""
try {
    if (Test-Path -LiteralPath $appDataRoot) {
        Remove-Item -LiteralPath $appDataRoot -Recurse -Force
    }
    $legacyVenv = Join-Path $PSScriptRoot ".venv"
    if (Test-Path -LiteralPath $legacyVenv) {
        Remove-Item -LiteralPath $legacyVenv -Recurse -Force
    }
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Installa_Slide_Narrator.ps1")
    exit $LASTEXITCODE
} catch {
    Write-Host "[ERRORE] $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Premere INVIO per chiudere"
    exit 1
}
