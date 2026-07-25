param([Parameter(Mandatory=$true)][string]$Manifest)
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot
$configPath = Join-Path $env:LOCALAPPDATA "SlideNarrator\runtime.json"
try {
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "Manifest non trovato: $Manifest" }
    if (-not (Test-Path -LiteralPath $configPath)) { throw "Runtime non installato. Eseguire Installa_Slide_Narrator.bat." }
    $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    if ($config.mode -eq "target") {
        $env:SLIDENARRATOR_SITE_PACKAGES = $config.site_packages
        & $config.python (Join-Path $PSScriptRoot "slide_narrator_bootstrap.py") (Join-Path $PSScriptRoot "slide_narrator_batch.py") $Manifest
    } else {
        & $config.python (Join-Path $PSScriptRoot "slide_narrator_batch.py") $Manifest
    }
    exit $LASTEXITCODE
} catch {
    Write-Host "[ERRORE] $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Premere INVIO per chiudere"
    exit 1
}
