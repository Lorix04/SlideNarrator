@echo off
setlocal EnableExtensions
cd /d "%~dp0"
if "%~1"=="" (
  echo Trascinare un manifest JSON su questo file oppure usare:
  echo AVVIA_BATCH.bat ESEMPIO_BATCH.json
  pause
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Avvia_Batch_Slide_Narrator.ps1" -Manifest "%~1"
exit /b %ERRORLEVEL%
