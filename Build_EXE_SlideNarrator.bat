@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

rem ====================================================================
rem  SlideNarrator - build Windows a falsi positivi ridotti
rem
rem  Output pubblico consigliato:
rem    dist\SlideNarrator\SlideNarrator.exe       (bundle ONEDIR)
rem    release\SlideNarrator_2.9.0_Windows_x64_portable.zip
rem
rem  Per compilare un bootloader PyInstaller univoco (richiede Visual
rem  Studio C++ Build Tools):
rem    Build_EXE_SlideNarrator.bat --custom-bootloader
rem ====================================================================

cd /d "%~dp0"
set "PROJECT_DIR=%CD%"
set "APP_NAME=SlideNarrator"
set "APP_VERSION=2.9.0"
set "ENTRY_FILE=slide_narrator_gui.py"
set "ICON_FILE=assets\slide_narrator_icon.ico"
set "VERSION_FILE=packaging\windows\SlideNarrator.version.txt"
set "MANIFEST_FILE=packaging\windows\SlideNarrator.manifest"
set "BUILD_VENV=.venv-build"
set "BUILD_PYTHON=%BUILD_VENV%\Scripts\python.exe"
set "OUTPUT_DIR=dist\%APP_NAME%"
set "OUTPUT_EXE=%OUTPUT_DIR%\%APP_NAME%.exe"
set "RELEASE_DIR=release"
set "RELEASE_ZIP=%RELEASE_DIR%\%APP_NAME%_%APP_VERSION%_Windows_x64_portable.zip"
set "RELEASE_HASH=%RELEASE_ZIP%.sha256.txt"
set "LOG_FILE=build_slidenarrator.log"
set "CUSTOM_BOOTLOADER=0"

if /I "%~1"=="--custom-bootloader" set "CUSTOM_BOOTLOADER=1"

>"%LOG_FILE%" echo [%date% %time%] Avvio build %APP_NAME% %APP_VERSION%

echo.
echo ====================================================================
echo   SlideNarrator - Build Windows verificabile
echo ====================================================================
echo Cartella progetto: %PROJECT_DIR%
echo Modalita':         ONEDIR, senza UPX
echo Versione:          %APP_VERSION%
if "%CUSTOM_BOOTLOADER%"=="1" (
    echo Bootloader:        compilato localmente da sorgente
) else (
    echo Bootloader:        PyInstaller ufficiale
)
echo.

rem --- Controlli preliminari ------------------------------------------
for %%F in ("%ENTRY_FILE%" "requirements.txt" "requirements-build.txt" "%ICON_FILE%" "%VERSION_FILE%" "%MANIFEST_FILE%") do (
    if not exist "%%~F" (
        echo [ERRORE] File richiesto non trovato: %%~F
        echo [ERRORE] File richiesto non trovato: %%~F>>"%LOG_FILE%"
        goto :fail
    )
)

rem --- Individuazione Python 3.11 x64 ---------------------------------
set "BASE_PYTHON="
where py >nul 2>&1
if not errorlevel 1 (
    py -3.11 -c "import struct,sys; assert sys.version_info[:2]==(3,11) and struct.calcsize('P')==8" >nul 2>&1
    if not errorlevel 1 set "BASE_PYTHON=py -3.11"
)
if not defined BASE_PYTHON (
    where python >nul 2>&1
    if not errorlevel 1 (
        python -c "import struct,sys; assert sys.version_info[:2]==(3,11) and struct.calcsize('P')==8" >nul 2>&1
        if not errorlevel 1 set "BASE_PYTHON=python"
    )
)
if not defined BASE_PYTHON (
    echo [ERRORE] Python 3.11 x64 non trovato sul PC di compilazione.
    echo Python serve soltanto per creare la release, non sul PC finale.
    goto :fail
)

echo Interprete build: %BASE_PYTHON%
echo Interprete build: %BASE_PYTHON%>>"%LOG_FILE%"

rem --- Ambiente pulito e riproducibile --------------------------------
echo.
echo [1/7] Creazione ambiente di build pulito...
if exist "%BUILD_VENV%" rmdir /s /q "%BUILD_VENV%"
%BASE_PYTHON% -m venv "%BUILD_VENV%" >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERRORE] Impossibile creare l'ambiente virtuale.
    goto :fail
)

set "PIP_DISABLE_PIP_VERSION_CHECK=1"
set "PIP_NO_CACHE_DIR=1"

rem --- Dipendenze minime ----------------------------------------------
echo.
echo [2/7] Installazione delle sole dipendenze dichiarate...
"%BUILD_PYTHON%" -m pip install --upgrade pip setuptools wheel >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto :deps_failed
"%BUILD_PYTHON%" -m pip install --no-cache-dir -r requirements.txt >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto :deps_failed

if "%CUSTOM_BOOTLOADER%"=="1" (
    echo Compilazione del bootloader PyInstaller da sorgente...
    set "PYINSTALLER_COMPILE_BOOTLOADER=1"
    "%BUILD_PYTHON%" -m pip install --no-cache-dir --verbose --no-binary=PyInstaller "PyInstaller==6.21.0" >>"%LOG_FILE%" 2>&1
    if errorlevel 1 (
        echo [ERRORE] Bootloader personalizzato non compilato.
        echo Installa Visual Studio C++ Build Tools e riprova.
        goto :fail
    )
) else (
    "%BUILD_PYTHON%" -m pip install --no-cache-dir -r requirements-build.txt >>"%LOG_FILE%" 2>&1
    if errorlevel 1 goto :deps_failed
)

"%BUILD_PYTHON%" -m pip check >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERRORE] pip check ha rilevato dipendenze incoerenti.
    goto :fail
)

rem --- Controllo sintassi ---------------------------------------------
echo.
echo [3/7] Controllo sintassi del progetto...
"%BUILD_PYTHON%" -m py_compile ^
    slide_narrator.py ^
    slide_narrator_gui.py ^
    slide_narrator_gui_legacy.py ^
    slide_narrator_batch.py ^
    slide_narrator_bootstrap.py ^
    voice_library.py ^
    voice_manager.py ^
    voice_clone.py ^
    video_export.py ^
    verifica_installazione.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERRORE] Uno o piu' file Python contengono errori di sintassi.
    goto :fail
)

rem --- Pulizia output precedenti --------------------------------------
echo.
echo [4/7] Pulizia degli artefatti precedenti...
if exist "build\pyinstaller" rmdir /s /q "build\pyinstaller"
if exist "dist" rmdir /s /q "dist"
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"
mkdir "%RELEASE_DIR%" >nul 2>&1

rem --- Build ONEDIR senza compressori ---------------------------------
echo.
echo [5/7] Generazione bundle ONEDIR senza UPX...
"%BUILD_PYTHON%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onedir ^
    --contents-directory "_internal" ^
    --noupx ^
    --windowed ^
    --name "%APP_NAME%" ^
    --workpath "build\pyinstaller" ^
    --specpath "build\pyinstaller" ^
    --distpath "dist" ^
    --icon "%ICON_FILE%" ^
    --version-file "%VERSION_FILE%" ^
    --manifest "%MANIFEST_FILE%" ^
    --add-data "assets;assets" ^
    --collect-all ttkbootstrap ^
    --collect-data edge_tts ^
    --collect-submodules edge_tts ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --exclude-module torch ^
    --exclude-module torchaudio ^
    --exclude-module torchvision ^
    --exclude-module tensorflow ^
    --exclude-module transformers ^
    --exclude-module TTS ^
    --exclude-module chatterbox ^
    --exclude-module pocket_tts ^
    --exclude-module sounddevice ^
    "%ENTRY_FILE%" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto :build_failed

if not exist "%OUTPUT_EXE%" (
    echo [ERRORE] PyInstaller non ha creato %OUTPUT_EXE%.
    goto :fail
)

rem --- Informazioni verificabili --------------------------------------
echo.
echo [6/7] Creazione metadati, ZIP e hash SHA-256...
(
    echo SlideNarrator %APP_VERSION%
    echo Build date: %date% %time%
    echo Python: 3.11 x64
    echo PyInstaller: 6.21.0
    echo Bundle: ONEDIR
    echo UPX: disabled
    if "%CUSTOM_BOOTLOADER%"=="1" echo Bootloader: locally compiled from source
    if "%CUSTOM_BOOTLOADER%"=="0" echo Bootloader: official PyInstaller distribution
) > "%OUTPUT_DIR%\BUILD_INFO.txt"

where git >nul 2>&1
if not errorlevel 1 (
    for /f %%C in ('git rev-parse HEAD 2^>nul') do echo Git commit: %%C>>"%OUTPUT_DIR%\BUILD_INFO.txt"
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Compress-Archive -Path '%OUTPUT_DIR%' -DestinationPath '%RELEASE_ZIP%' -CompressionLevel Optimal -Force; $h=(Get-FileHash '%RELEASE_ZIP%' -Algorithm SHA256).Hash.ToLower(); Set-Content -Encoding ASCII '%RELEASE_HASH%' ($h + '  ' + [IO.Path]::GetFileName('%RELEASE_ZIP%'))" >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERRORE] Creazione ZIP o hash non riuscita.
    goto :fail
)

rem --- Verifica locale -------------------------------------------------
echo.
echo [7/7] Verifica release e scansione Defender, se disponibile...
powershell -NoProfile -ExecutionPolicy Bypass -File "packaging\windows\Verifica_RELEASE_SlideNarrator.ps1" -ProjectRoot "%PROJECT_DIR%" >>"%LOG_FILE%" 2>&1
set "VERIFY_EXIT=%ERRORLEVEL%"
if not "%VERIFY_EXIT%"=="0" (
    echo [ATTENZIONE] La verifica release ha restituito codice %VERIFY_EXIT%.
    echo Controlla il log prima di pubblicare.
)

echo.
echo ====================================================================
echo   BUILD COMPLETATA
echo ====================================================================
echo Cartella portabile:
echo   %PROJECT_DIR%\%OUTPUT_DIR%
echo.
echo Archivio da pubblicare:
echo   %PROJECT_DIR%\%RELEASE_ZIP%
echo.
echo Hash:
echo   %PROJECT_DIR%\%RELEASE_HASH%
echo.
echo Log:
echo   %PROJECT_DIR%\%LOG_FILE%
echo.
echo IMPORTANTE: non pubblicare una build non firmata come release stabile.
echo Firma l'EXE e l'installer con Firma_RELEASE_SlideNarrator.bat,
echo poi ripeti la scansione e invia eventuali falsi positivi ai vendor.
echo.
choice /C SN /N /M "Avviare il programma dalla cartella portabile? [S/N]: "
if errorlevel 2 goto :done
start "" "%OUTPUT_EXE%"
goto :done

:deps_failed
echo [ERRORE] Installazione delle dipendenze non riuscita.
goto :fail

:build_failed
echo [ERRORE] La compilazione PyInstaller non e' riuscita.
echo Controlla: %PROJECT_DIR%\%LOG_FILE%
goto :fail

:fail
echo.
echo Operazione non completata.
echo Consulta il log: %PROJECT_DIR%\%LOG_FILE%
echo.
pause
exit /b 1

:done
echo.
echo Operazione completata.
pause
exit /b 0

