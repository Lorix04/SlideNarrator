@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

if not exist "dist\SlideNarrator\SlideNarrator.exe" (
    echo [ERRORE] Prima esegui Build_EXE_SlideNarrator.bat
    pause
    exit /b 1
)

set "ISCC="
for /f "delims=" %%F in ('where ISCC.exe 2^>nul') do if not defined ISCC set "ISCC=%%F"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 7\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"

if not defined ISCC (
    echo [ERRORE] Inno Setup non trovato.
    echo Installa Inno Setup, poi riprova.
    pause
    exit /b 1
)

if exist "installer\output" rmdir /s /q "installer\output"
"%ISCC%" "installer\SlideNarrator.iss"
if errorlevel 1 (
    echo [ERRORE] Creazione installer non riuscita.
    pause
    exit /b 1
)

echo.
echo Installer creato in installer\output\
echo Per una release pubblica firmalo con Firma_RELEASE_SlideNarrator.bat.
pause
exit /b 0
