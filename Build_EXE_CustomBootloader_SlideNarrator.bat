@echo off
setlocal
cd /d "%~dp0"
echo Questa modalita' compila un bootloader PyInstaller univoco.
echo Richiede Visual Studio C++ Build Tools installato.
echo.
call "Build_EXE_SlideNarrator.bat" --custom-bootloader
exit /b %ERRORLEVEL%
