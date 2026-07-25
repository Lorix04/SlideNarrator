@echo off
setlocal
cd /d "%~dp0"
if not exist "dist\SlideNarrator\SlideNarrator.exe" (
  echo [ERRORE] Prima genera la release con Build_EXE_SlideNarrator.bat
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "packaging\windows\Verifica_RELEASE_SlideNarrator.ps1" -ProjectRoot "%CD%"
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" echo Verifica terminata con codice %RC%.
pause
exit /b %RC%

