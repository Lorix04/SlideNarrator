@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Slide Narrator - Rimozione componenti installati

echo.
echo ====================================================================
echo   Slide Narrator - Rimozione componenti installati
echo ====================================================================
echo.
echo Questa procedura rimuove:
echo   - il runtime isolato in %%LOCALAPPDATA%%\SlideNarrator
echo   - tutte le librerie Python installate per Slide Narrator
echo   - PocketTTS, Chatterbox e Coqui XTTS, se presenti nel runtime
echo   - gli ambienti virtuali locali legacy (.venv)
echo   - cache Python e cache locali del progetto
echo.
echo NON rimuove:
echo   - Python
echo   - pip dell'installazione principale di Python
echo   - i file sorgente di Slide Narrator
echo   - presentazioni, Excel, audio, video o altri documenti
echo.
echo FFmpeg e LibreOffice verranno gestiti separatamente:
echo saranno rimossi solo se confermi esplicitamente.
echo.

choice /C SN /N /M "Continuare con la pulizia di Slide Narrator? [S/N]: "
if errorlevel 2 (
    echo.
    echo Operazione annullata.
    goto :END_OK
)

set "REMOVE_ERRORS=0"
set "APP_ROOT=%LOCALAPPDATA%\SlideNarrator"
set "LEGACY_VENV=%~dp0.venv"

echo.
echo --------------------------------------------------------------------
echo Rimozione runtime e librerie di Slide Narrator
echo --------------------------------------------------------------------

if exist "%APP_ROOT%" (
    echo Rimuovo: "%APP_ROOT%"
    rmdir /S /Q "%APP_ROOT%" 2>nul
    if exist "%APP_ROOT%" (
        echo [ERRORE] Non e' stato possibile eliminare "%APP_ROOT%".
        echo Chiudi Slide Narrator e ogni finestra Python che lo sta usando.
        set "REMOVE_ERRORS=1"
    ) else (
        echo [OK] Runtime e librerie rimossi.
    )
) else (
    echo [INFO] Runtime non presente.
)

if exist "%LEGACY_VENV%" (
    echo Rimuovo ambiente virtuale locale: "%LEGACY_VENV%"
    rmdir /S /Q "%LEGACY_VENV%" 2>nul
    if exist "%LEGACY_VENV%" (
        echo [ERRORE] Non e' stato possibile eliminare "%LEGACY_VENV%".
        set "REMOVE_ERRORS=1"
    ) else (
        echo [OK] Ambiente virtuale locale rimosso.
    )
)

echo.
echo --------------------------------------------------------------------
echo Rimozione cache locali del progetto
echo --------------------------------------------------------------------

for /D /R "%~dp0" %%D in (__pycache__) do (
    if exist "%%~fD" rmdir /S /Q "%%~fD" 2>nul
)

for /D /R "%~dp0" %%D in (.slide_narrator_cache) do (
    if exist "%%~fD" rmdir /S /Q "%%~fD" 2>nul
)

if exist "%~dp0installazione_slide_narrator.log" del /F /Q "%~dp0installazione_slide_narrator.log" 2>nul
if exist "%~dp0verifica_installazione_finale.txt" del /F /Q "%~dp0verifica_installazione_finale.txt" 2>nul

echo [OK] Pulizia delle cache completata.

where winget.exe >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] WinGet non e' disponibile: FFmpeg e LibreOffice non vengono modificati.
    goto :SUMMARY
)

echo.
echo --------------------------------------------------------------------
echo Componenti di sistema opzionali
echo --------------------------------------------------------------------
echo.
echo ATTENZIONE: FFmpeg puo' essere usato anche da altri programmi.
choice /C SN /N /M "Disinstallare FFmpeg dal computer? [S/N]: "
if errorlevel 2 goto :ASK_LIBREOFFICE

echo.
echo Avvio disinstallazione di FFmpeg...
winget uninstall --id Gyan.FFmpeg -e --source winget
if errorlevel 1 (
    echo [AVVISO] FFmpeg non e' stato disinstallato oppure non era presente.
    set "REMOVE_ERRORS=1"
) else (
    echo [OK] Comando di disinstallazione FFmpeg completato.
)

:ASK_LIBREOFFICE
echo.
echo ATTENZIONE: LibreOffice e' un programma indipendente e puo' servire
echo anche per aprire documenti al di fuori di Slide Narrator.
choice /C SN /N /M "Disinstallare LibreOffice dal computer? [S/N]: "
if errorlevel 2 goto :SUMMARY

echo.
echo Avvio disinstallazione di LibreOffice...
winget uninstall --id TheDocumentFoundation.LibreOffice -e --source winget
if errorlevel 1 (
    echo [AVVISO] LibreOffice non e' stato disinstallato oppure non era presente.
    set "REMOVE_ERRORS=1"
) else (
    echo [OK] Comando di disinstallazione LibreOffice completato.
)

:SUMMARY
echo.
echo ====================================================================
echo   Operazione completata
echo ====================================================================
echo.
echo Python NON e' stato rimosso o modificato.
echo I file del progetto e i documenti personali NON sono stati eliminati.
echo.
if "%REMOVE_ERRORS%"=="1" (
    echo Alcune operazioni non sono riuscite completamente.
    echo Chiudi Slide Narrator e riprova, eventualmente come amministratore.
    goto :END_ERROR
)

echo Tutti i componenti selezionati sono stati rimossi.
goto :END_OK

:END_ERROR
echo.
pause
exit /b 1

:END_OK
echo.
pause
exit /b 0
