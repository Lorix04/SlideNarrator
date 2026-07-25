@echo off
setlocal EnableExtensions
cd /d "%~dp0"
where powershell.exe >nul 2>&1
if errorlevel 1 (
  echo [ERRORE] Windows PowerShell non e' disponibile.
  pause
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installa_Slide_Narrator.ps1"
exit /b %ERRORLEVEL%
