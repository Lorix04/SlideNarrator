$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot
$configPath = Join-Path $env:LOCALAPPDATA "SlideNarrator\runtime.json"
try {
    if (-not (Test-Path -LiteralPath $configPath)) {
        throw "Runtime non installato. Eseguire prima Installa_Slide_Narrator.bat."
    }
    $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    if (-not (Test-Path -LiteralPath $config.python)) {
        throw "Interprete del runtime non trovato: $($config.python). Eseguire Ripara_Installazione_Slide_Narrator.bat."
    }
    if ($config.mode -eq "target") {
        if (-not (Test-Path -LiteralPath $config.site_packages)) {
            throw "Pacchetti del runtime non trovati. Eseguire Ripara_Installazione_Slide_Narrator.bat."
        }
        $env:SLIDENARRATOR_SITE_PACKAGES = $config.site_packages
        & $config.python (Join-Path $PSScriptRoot "slide_narrator_bootstrap.py") (Join-Path $PSScriptRoot "slide_narrator_gui.py")
    } else {
        & $config.python (Join-Path $PSScriptRoot "slide_narrator_gui.py")
    }
    exit $LASTEXITCODE
} catch {
    Write-Host "[ERRORE] $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Premere INVIO per chiudere"
    exit 1
}
