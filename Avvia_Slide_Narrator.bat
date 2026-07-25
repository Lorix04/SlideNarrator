@echo off
setlocal EnableExtensions
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Avvia_Slide_Narrator.ps1"
exit /b %ERRORLEVEL%
