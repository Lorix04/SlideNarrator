param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"
$project = (Resolve-Path $ProjectRoot).Path
$exe = Join-Path $project "dist\SlideNarrator\SlideNarrator.exe"
$zip = Join-Path $project "release\SlideNarrator_2.9.0_Windows_x64_portable.zip"
$report = Join-Path $project "release\verifica_release.txt"
$lines = [System.Collections.Generic.List[string]]::new()

function Add-Result([string]$Name, [bool]$Ok, [string]$Detail) {
    $status = if ($Ok) { "PASS" } else { "FAIL" }
    $lines.Add("[$status] $Name - $Detail")
}

Add-Result "EXE presente" (Test-Path $exe) $exe
Add-Result "ZIP presente" (Test-Path $zip) $zip
if (-not (Test-Path $exe) -or -not (Test-Path $zip)) {
    $lines | Set-Content -Encoding UTF8 $report
    $lines | ForEach-Object { Write-Host $_ }
    exit 10
}

$exeHash = (Get-FileHash $exe -Algorithm SHA256).Hash.ToLowerInvariant()
$zipHash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLowerInvariant()
$lines.Add("SHA256 EXE: $exeHash")
$lines.Add("SHA256 ZIP: $zipHash")

$sig = Get-AuthenticodeSignature $exe
$signatureOk = $sig.Status -eq "Valid"
Add-Result "Firma Authenticode" $signatureOk ("Stato: " + $sig.Status)
if (-not $signatureOk) {
    $lines.Add("[INFO] La build e' utilizzabile per test, ma non va pubblicata come release stabile senza firma valida.")
}

$defender = $null
$platformRoot = Join-Path $env:ProgramData "Microsoft\Windows Defender\Platform"
if (Test-Path $platformRoot) {
    $candidate = Get-ChildItem $platformRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
    if ($candidate) {
        $path = Join-Path $candidate.FullName "MpCmdRun.exe"
        if (Test-Path $path) { $defender = $path }
    }
}
if (-not $defender) {
    $fallback = Join-Path $env:ProgramFiles "Windows Defender\MpCmdRun.exe"
    if (Test-Path $fallback) { $defender = $fallback }
}

$defenderOk = $true
if ($defender) {
    & $defender -Scan -ScanType 3 -File (Join-Path $project "dist\SlideNarrator") -DisableRemediation
    $scanCode = $LASTEXITCODE
    $defenderOk = $scanCode -eq 0
    Add-Result "Microsoft Defender" $defenderOk "Codice scansione: $scanCode"
} else {
    $lines.Add("[SKIP] Microsoft Defender - MpCmdRun.exe non trovato")
}

$lines | Set-Content -Encoding UTF8 $report
$lines | ForEach-Object { Write-Host $_ }
if (-not $defenderOk) { exit 20 }
exit 0
