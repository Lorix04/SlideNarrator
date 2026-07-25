@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "APP_EXE=dist\SlideNarrator\SlideNarrator.exe"
set "SETUP_EXE=installer\output\SlideNarrator_Setup_2.9.0.exe"
set "RELEASE_ZIP=release\SlideNarrator_2.9.0_Windows_x64_portable.zip"
set "RELEASE_HASH=%RELEASE_ZIP%.sha256.txt"
set "TIMESTAMP_URL=http://timestamp.digicert.com"

if not defined SLIDENARRATOR_CERT_SUBJECT (
    echo [ERRORE] Imposta prima il nome del certificato Authenticode:
    echo   set "SLIDENARRATOR_CERT_SUBJECT=Nome esatto dell'editore"
    echo.
    echo Il certificato deve essere presente nell'archivio certificati Windows.
    pause
    exit /b 1
)

set "SIGNTOOL="
for /f "delims=" %%F in ('where signtool.exe 2^>nul') do if not defined SIGNTOOL set "SIGNTOOL=%%F"
if not defined SIGNTOOL (
    for /f "delims=" %%F in ('dir "%ProgramFiles(x86)%\Windows Kits\10\bin\*\x64\signtool.exe" /b /s /o:-n 2^>nul') do if not defined SIGNTOOL set "SIGNTOOL=%%F"
)
if not defined SIGNTOOL (
    echo [ERRORE] signtool.exe non trovato. Installa Windows SDK.
    pause
    exit /b 1
)

if not exist "%APP_EXE%" (
    echo [ERRORE] File non trovato: %APP_EXE%
    echo Esegui prima Build_EXE_SlideNarrator.bat
    pause
    exit /b 1
)

echo Firma applicazione...
"%SIGNTOOL%" sign /a /n "%SLIDENARRATOR_CERT_SUBJECT%" /fd SHA256 /tr "%TIMESTAMP_URL%" /td SHA256 /d "SlideNarrator" /du "https://github.com/Lorix04/SlideNarrator" "%APP_EXE%"
if errorlevel 1 goto :fail
"%SIGNTOOL%" verify /pa /v "%APP_EXE%"
if errorlevel 1 goto :fail

if exist "%SETUP_EXE%" (
    echo Firma installer...
    "%SIGNTOOL%" sign /a /n "%SLIDENARRATOR_CERT_SUBJECT%" /fd SHA256 /tr "%TIMESTAMP_URL%" /td SHA256 /d "SlideNarrator Setup" /du "https://github.com/Lorix04/SlideNarrator" "%SETUP_EXE%"
    if errorlevel 1 goto :fail
    "%SIGNTOOL%" verify /pa /v "%SETUP_EXE%"
    if errorlevel 1 goto :fail
)

echo Rigenerazione archivio portabile e SHA-256 dopo la firma...
if not exist "release" mkdir "release"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\SlideNarrator' -DestinationPath '%RELEASE_ZIP%' -CompressionLevel Optimal -Force; $h=(Get-FileHash '%RELEASE_ZIP%' -Algorithm SHA256).Hash.ToLower(); Set-Content -Encoding ASCII '%RELEASE_HASH%' ($h + '  ' + [IO.Path]::GetFileName('%RELEASE_ZIP%'))"
if errorlevel 1 goto :fail

echo.
echo Firma, verifica e aggiornamento hash completati.
pause
exit /b 0

:fail
echo.
echo [ERRORE] Firma o verifica non riuscita.
pause
exit /b 1

