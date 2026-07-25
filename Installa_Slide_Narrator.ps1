$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$ProjectLogFile = Join-Path $PSScriptRoot "installazione_slide_narrator.log"
$LogFile = $null
$TranscriptStarted = $false
try {
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "SlideNarrator"
    [System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile = Join-Path $tempRoot ("installazione_slide_narrator_{0}_{1}.log" -f $stamp, $PID)
    Start-Transcript -Path $LogFile -Force | Out-Null
    $TranscriptStarted = $true
} catch {
    $LogFile = $null
}

function Write-Section([string]$Text) {
    Write-Host ""
    Write-Host ("=" * 68) -ForegroundColor DarkCyan
    Write-Host ("  " + $Text) -ForegroundColor Cyan
    Write-Host ("=" * 68) -ForegroundColor DarkCyan
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [Parameter(Mandatory=$true)][string]$Description
    )
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Description non riuscita. Codice di uscita: $LASTEXITCODE"
    }
}

function Ask-YesNo {
    param(
        [Parameter(Mandatory=$true)][string]$Question,
        [bool]$DefaultYes = $false
    )
    $suffix = if ($DefaultYes) { "[S/n]" } else { "[s/N]" }
    $answer = (Read-Host "$Question $suffix").Trim()
    if ([string]::IsNullOrWhiteSpace($answer)) { return $DefaultYes }
    return $answer -match '^(s|si|si|y|yes)$'
}

function Find-CompatiblePython {
    $candidates = @(
        @{ Command = "python.exe"; Args = @() },
        @{ Command = "python3.exe"; Args = @() },
        @{ Command = "py.exe"; Args = @("-3.14") },
        @{ Command = "py.exe"; Args = @("-3.13") },
        @{ Command = "py.exe"; Args = @("-3.12") },
        @{ Command = "py.exe"; Args = @("-3.11") },
        @{ Command = "py.exe"; Args = @("-3.10") },
        @{ Command = "py.exe"; Args = @("-3") }
    )

    foreach ($candidate in $candidates) {
        $resolved = Get-Command $candidate.Command -ErrorAction SilentlyContinue
        if (-not $resolved) { continue }
        $probe = @($candidate.Args) + @(
            "-c",
            "import sys; ok=(3,10)<=sys.version_info[:2]<=(3,14); print(sys.executable if ok else ''); raise SystemExit(0 if ok else 1)"
        )
        $previousErrorAction = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $out = & $resolved.Source @probe 2>$null
            $probeExitCode = $LASTEXITCODE
        } catch {
            $probeExitCode = 1
            $out = $null
        } finally {
            $ErrorActionPreference = $previousErrorAction
        }
        if ($probeExitCode -ne 0 -or -not $out) { continue }
        $exe = ($out | Select-Object -Last 1).ToString().Trim()
        if ($exe -and (Test-Path -LiteralPath $exe)) {
            return (Resolve-Path -LiteralPath $exe).Path
        }
    }
    return $null
}

function Test-PythonCommand {
    param([string]$Python, [string[]]$PrefixArgs = @())
    try {
        & $Python @PrefixArgs -c "import sys; print(sys.executable)" *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Ensure-BasePip {
    param([string]$Python)
    & $Python -m pip --version *> $null
    if ($LASTEXITCODE -eq 0) { return }
    Write-Host "pip non e' disponibile nell'interprete base. Eseguo ensurepip..." -ForegroundColor Yellow
    Invoke-Checked -FilePath $Python -Arguments @("-m", "ensurepip", "--upgrade") -Description "Installazione di pip con ensurepip"
}

function Invoke-TargetRuntime {
    param(
        [string]$Python,
        [string]$SitePackages,
        [string[]]$Arguments
    )
    $old = $env:SLIDENARRATOR_SITE_PACKAGES
    try {
        $env:SLIDENARRATOR_SITE_PACKAGES = $SitePackages
        & $Python (Join-Path $PSScriptRoot "slide_narrator_bootstrap.py") @Arguments | Out-Host
        $code = $LASTEXITCODE
        return [int]$code
    } finally {
        $env:SLIDENARRATOR_SITE_PACKAGES = $old
    }
}

function Save-RuntimeConfig {
    param(
        [string]$ConfigPath,
        [string]$Mode,
        [string]$Python,
        [string]$RuntimeRoot,
        [string]$SitePackages,
        [string]$Version
    )
    $configDir = Split-Path -Parent $ConfigPath
    [System.IO.Directory]::CreateDirectory($configDir) | Out-Null
    $obj = [ordered]@{
        schema = 2
        mode = $Mode
        python = $Python
        runtime_root = $RuntimeRoot
        site_packages = $SitePackages
        version = $Version
        updated_at = (Get-Date).ToString("o")
    }
    $json = $obj | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($ConfigPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

$exitCode = 1
$failureMessage = $null
try {
    Write-Section "Slide Narrator - Installazione Windows"
    Write-Host "Cartella progetto: $PSScriptRoot"

    $pythonExe = Find-CompatiblePython
    if (-not $pythonExe) {
        throw "Serve Python da 3.10 a 3.14 a 64 bit."
    }

    $pythonVersion = & $pythonExe --version 2>&1
    if ($LASTEXITCODE -ne 0) { throw "L'interprete Python rilevato non puo' essere eseguito." }
    $versionTag = & $pythonExe -c "import sys; print(f'py{sys.version_info.major}{sys.version_info.minor}')"
    $versionPlain = & $pythonExe -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    if ($LASTEXITCODE -ne 0 -or -not $versionTag) { throw "Impossibile leggere la versione Python." }

    Write-Host "Interprete: $pythonVersion" -ForegroundColor Green
    Write-Host "Percorso:    $pythonExe"

    # L'ambiente non viene piu' creato dentro Documenti. Windows Defender e
    # Controlled Folder Access possono bloccare i launcher .exe dei venv nelle
    # cartelle protette. LocalAppData e' il percorso corretto per il runtime.
    $appDataRoot = Join-Path $env:LOCALAPPDATA "SlideNarrator"
    $runtimeRoot = Join-Path $appDataRoot ("runtime\" + $versionTag)
    $venvDir = Join-Path $runtimeRoot "venv"
    $venvPython = Join-Path $venvDir "Scripts\python.exe"
    $sitePackages = Join-Path $runtimeRoot "site-packages"
    $configPath = Join-Path $appDataRoot "runtime.json"
    [System.IO.Directory]::CreateDirectory($runtimeRoot) | Out-Null

    Write-Host "Runtime:     $runtimeRoot"
    $runtimeMode = $null
    $runtimePython = $null

    if (Test-Path -LiteralPath $venvPython) {
        if (Test-PythonCommand -Python $venvPython) {
            Write-Host "Ambiente virtuale gia' presente: lo riutilizzo." -ForegroundColor Green
            $runtimeMode = "venv"
            $runtimePython = $venvPython
        } else {
            Write-Host "Ambiente virtuale incompleto: lo rimuovo." -ForegroundColor Yellow
            Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not $runtimeMode) {
        Write-Host "Creo l'ambiente virtuale in LocalAppData..."
        if (Test-Path -LiteralPath $venvDir) {
            Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        $previousErrorAction = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            & $pythonExe -m venv $venvDir
            $venvExit = $LASTEXITCODE
        } catch {
            $venvExit = 1
        } finally {
            $ErrorActionPreference = $previousErrorAction
        }

        if ($venvExit -eq 0 -and (Test-Path -LiteralPath $venvPython) -and (Test-PythonCommand -Python $venvPython)) {
            $runtimeMode = "venv"
            $runtimePython = $venvPython
        } else {
            Write-Host "Il modulo venv non e' utilizzabile su questo PC." -ForegroundColor Yellow
            Write-Host "Attivo il runtime isolato alternativo in LocalAppData..." -ForegroundColor Yellow
            Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
            Ensure-BasePip -Python $pythonExe
            if (Test-Path -LiteralPath $sitePackages) {
                Remove-Item -LiteralPath $sitePackages -Recurse -Force -ErrorAction SilentlyContinue
            }
            [System.IO.Directory]::CreateDirectory($sitePackages) | Out-Null
            Invoke-Checked -FilePath $pythonExe -Arguments @(
                "-m", "pip", "install", "--upgrade", "--prefer-binary",
                "--no-warn-script-location", "--target", $sitePackages,
                "-r", (Join-Path $PSScriptRoot "requirements.txt")
            ) -Description "Installazione del runtime isolato"
            $runtimeMode = "target"
            $runtimePython = $pythonExe
        }
    }

    $oldVerifyOutput = $env:SLIDENARRATOR_VERIFY_OUTPUT
    $env:SLIDENARRATOR_VERIFY_OUTPUT = Join-Path $runtimeRoot "verifica_installazione.json"

    Write-Section "Dipendenze base"
    if ($runtimeMode -eq "venv") {
        & $runtimePython -m pip --version *> $null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "pip non e' presente nell'ambiente. Eseguo ensurepip..."
            Invoke-Checked -FilePath $runtimePython -Arguments @("-m", "ensurepip", "--upgrade") -Description "Installazione di pip con ensurepip"
        }
        Invoke-Checked -FilePath $runtimePython -Arguments @("-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel") -Description "Aggiornamento degli strumenti Python"
        Invoke-Checked -FilePath $runtimePython -Arguments @("-m", "pip", "install", "--prefer-binary", "-r", (Join-Path $PSScriptRoot "requirements.txt")) -Description "Installazione delle dipendenze base"
        Invoke-Checked -FilePath $runtimePython -Arguments @("-m", "pip", "check") -Description "Controllo delle dipendenze"
        Invoke-Checked -FilePath $runtimePython -Arguments @((Join-Path $PSScriptRoot "verifica_installazione.py")) -Description "Verifica del progetto"
    } else {
        $rc = Invoke-TargetRuntime -Python $runtimePython -SitePackages $sitePackages -Arguments @((Join-Path $PSScriptRoot "verifica_installazione.py"))
        if ($rc -ne 0) { throw "Verifica del runtime isolato non riuscita. Codice di uscita: $rc" }
    }

    Save-RuntimeConfig -ConfigPath $configPath -Mode $runtimeMode -Python $runtimePython -RuntimeRoot $runtimeRoot -SitePackages $sitePackages -Version $versionPlain

    $cloneStatus = "non richiesti"
    if (Ask-YesNo -Question "Installare PocketTTS e Chatterbox? Richiedono diversi GB" -DefaultYes $false) {
        if ($runtimeMode -eq "venv") {
            & $runtimePython -m pip install --prefer-binary -r (Join-Path $PSScriptRoot "requirements-clone.txt")
        } else {
            & $runtimePython -m pip install --upgrade --prefer-binary --no-warn-script-location --target $sitePackages -r (Join-Path $PSScriptRoot "requirements-clone.txt")
        }
        $cloneStatus = if ($LASTEXITCODE -eq 0) { "OK" } else { "FALLITI o incompleti" }
    }

    $xttsStatus = "non richiesto"
    if (Ask-YesNo -Question "Installare anche Coqui XTTS? Verificare la licenza del modello" -DefaultYes $false) {
        if ($runtimeMode -eq "venv") {
            & $runtimePython -m pip install --prefer-binary -r (Join-Path $PSScriptRoot "requirements-xtts.txt")
        } else {
            & $runtimePython -m pip install --upgrade --prefer-binary --no-warn-script-location --target $sitePackages -r (Join-Path $PSScriptRoot "requirements-xtts.txt")
        }
        $xttsStatus = if ($LASTEXITCODE -eq 0) { "OK" } else { "FALLITO" }
    }

    Write-Section "Strumenti video opzionali"
    $ffmpegStatus = if (Get-Command ffmpeg -ErrorAction SilentlyContinue) { "OK" } else { "NON TROVATO" }
    if ($ffmpegStatus -ne "OK" -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        if (Ask-YesNo -Question "Installare FFmpeg con winget" -DefaultYes $true) {
            & winget install --id Gyan.FFmpeg -e --source winget --accept-package-agreements --accept-source-agreements
            $ffmpegStatus = if ($LASTEXITCODE -eq 0) { "installato; riaprire il programma se non viene rilevato subito" } else { "installazione fallita" }
        }
    }

    $soffice = Get-Command soffice -ErrorAction SilentlyContinue
    $libreOfficePath = "C:\Program Files\LibreOffice\program\soffice.exe"
    $libreStatus = if ($soffice -or (Test-Path -LiteralPath $libreOfficePath)) { "OK" } else { "NON TROVATO" }
    if ($libreStatus -ne "OK" -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        if (Ask-YesNo -Question "Installare LibreOffice per l'output video" -DefaultYes $true) {
            & winget install --id TheDocumentFoundation.LibreOffice -e --source winget --accept-package-agreements --accept-source-agreements
            $libreStatus = if ($LASTEXITCODE -eq 0) { "installato" } else { "installazione fallita" }
        }
    }

    $verifyFile = Join-Path $runtimeRoot "verifica_installazione_finale.txt"
    # Alcune librerie opzionali (per esempio Transformers senza PyTorch)
    # scrivono semplici avvisi su stderr pur terminando con codice 0.
    # Con ErrorActionPreference=Stop PowerShell 5.1 puo' interpretarli come
    # errori terminanti durante il reindirizzamento. La verifica deve quindi
    # basarsi sul codice di uscita del processo Python, non sul solo stderr.
    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        if ($runtimeMode -eq "venv") {
            & $runtimePython (Join-Path $PSScriptRoot "verifica_installazione.py") *> $verifyFile
            $verifyExit = $LASTEXITCODE
        } else {
            $oldSite = $env:SLIDENARRATOR_SITE_PACKAGES
            try {
                $env:SLIDENARRATOR_SITE_PACKAGES = $sitePackages
                & $runtimePython (Join-Path $PSScriptRoot "slide_narrator_bootstrap.py") (Join-Path $PSScriptRoot "verifica_installazione.py") *> $verifyFile
                $verifyExit = $LASTEXITCODE
            } finally {
                $env:SLIDENARRATOR_SITE_PACKAGES = $oldSite
            }
        }
    } finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($verifyExit -ne 0) {
        throw "La verifica finale non e' riuscita. Consultare verifica_installazione_finale.txt."
    }

    Write-Section "Installazione completata"
    Write-Host "Base:         OK" -ForegroundColor Green
    Write-Host "Runtime:      $runtimeMode in $runtimeRoot"
    Write-Host "Voci locali:  $cloneStatus"
    Write-Host "XTTS:         $xttsStatus"
    Write-Host "FFmpeg:       $ffmpegStatus"
    Write-Host "LibreOffice:  $libreStatus"
    Write-Host ""
    Write-Host "Avviare il programma con Avvia_Slide_Narrator.bat" -ForegroundColor Cyan
    $exitCode = 0
} catch {
    $failureMessage = $_.Exception.Message
    $exitCode = 1
    Write-Host ""
    Write-Host "[ERRORE] $failureMessage" -ForegroundColor Red
} finally {
    if (Get-Variable -Name oldVerifyOutput -ErrorAction SilentlyContinue) {
        $env:SLIDENARRATOR_VERIFY_OUTPUT = $oldVerifyOutput
    }
    if ($TranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
    }
}

$availableLog = $null
if ($LogFile -and (Test-Path -LiteralPath $LogFile)) {
    $availableLog = $LogFile
    try {
        Copy-Item -LiteralPath $LogFile -Destination $ProjectLogFile -Force -ErrorAction Stop
        $availableLog = $ProjectLogFile
    } catch {}
}

if ($exitCode -ne 0) {
    Write-Host ""
    Write-Host "[ERRORE] $failureMessage" -ForegroundColor Red
}
if ($availableLog) {
    Write-Host "Log installazione: $availableLog" -ForegroundColor Yellow
} else {
    Write-Host "[AVVISO] Non e' stato possibile creare il log di installazione." -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Premere INVIO per chiudere"
exit $exitCode
