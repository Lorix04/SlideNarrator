@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   PPTX TTS  -  Installazione
echo ============================================================
echo.

REM --- 1. Cerco un interprete Python adatto (3.10 - 3.14) ------
set "PYLAUNCH="
py -3.14 -c "import sys" >nul 2>&1 && set "PYLAUNCH=py -3.14"
if not defined PYLAUNCH (
  py -3 -c "import sys" >nul 2>&1 && set "PYLAUNCH=py -3"
)
if not defined PYLAUNCH (
  python -c "import sys" >nul 2>&1 && set "PYLAUNCH=python"
)

if not defined PYLAUNCH (
  echo [ERRORE] Python non trovato.
  echo Installa Python 3.10 - 3.14 da https://www.python.org/downloads/
  echo Durante l'installazione spunta "Add Python to PATH".
  echo.
  pause
  exit /b 1
)
echo Interprete Python: %PYLAUNCH%
for /f "delims=" %%v in ('%PYLAUNCH% --version') do echo   %%v
echo.

REM --- 2. Creo l'ambiente virtuale .venv ----------------------
if exist ".venv\Scripts\python.exe" (
  echo Ambiente virtuale .venv gia' presente, lo riuso.
) else (
  echo Creo l'ambiente virtuale .venv ...
  %PYLAUNCH% -m venv .venv
  if errorlevel 1 (
    echo [ERRORE] Creazione dell'ambiente virtuale fallita.
    pause
    exit /b 1
  )
)
set "VPY=.venv\Scripts\python.exe"
echo.

REM --- 3. Aggiorno pip e installo i pacchetti base ------------
echo Aggiorno pip ...
"%VPY%" -m pip install --upgrade pip >nul

echo Installo i pacchetti base ^(interfaccia, voci Microsoft, supporto video^) ...
"%VPY%" -m pip install edge-tts python-pptx openpyxl lxml mutagen pygame-ce pymupdf
if errorlevel 1 (
  echo [ERRORE] Installazione dei pacchetti base fallita.
  pause
  exit /b 1
)
echo.

REM --- 4. Motori per le voci clonate (opzionali, pesanti) -----
echo ------------------------------------------------------------
echo   Voci clonate  ^(PocketTTS + Chatterbox^)
echo   Questi motori scaricano PyTorch e i modelli: diversi GB,
echo   puo' volerci parecchio. Se non ti servono subito, rispondi N.
echo ------------------------------------------------------------
set /p CLONE="Installare i motori per le voci clonate? [S/n]: "
if /i "!CLONE!"=="n" (
  echo Salto i motori di cloning. Potrai installarli dopo rilanciando questo file.
) else (
  echo Installo pocket-tts e sounddevice ...
  "%VPY%" -m pip install pocket-tts sounddevice
  echo Installo chatterbox-tts ...
  "%VPY%" -m pip install chatterbox-tts
  echo.
  echo NOTA: per clonare la tua voce con PocketTTS serve l'accesso al modello
  echo riservato 'kyutai/pocket-tts' su Hugging Face ^(gratis, ma va richiesto^).
  set /p HFLOGIN="Vuoi autenticarti ora su Hugging Face? [S/n]: "
  if /i not "!HFLOGIN!"=="n" (
    if exist ".venv\Scripts\hf.exe" (
      ".venv\Scripts\hf.exe" auth login
    ) else (
      if exist ".venv\Scripts\huggingface-cli.exe" (
        echo Il comando classico e' deprecato; provo comunque...
        ".venv\Scripts\huggingface-cli.exe" login
      ) else (
        echo Comando 'hf' non trovato. Autenticati a mano con:  .venv\Scripts\hf auth login
      )
    )
  )
)
echo.

REM --- 4b. XTTS v2 (opzionale, SOLO USO NON COMMERCIALE) -------
echo ------------------------------------------------------------
echo   XTTS v2  ^(qualita' alta, italiano - SOLO USO NON COMMERCIALE^)
echo   Licenza Coqui CPML: NON puo' essere usato per materiale
echo   venduto o per i corsi a pagamento. Su CPU e' lento
echo   ^(~10-30s a frase^): adatto a batch lasciati girare.
echo   Per i corsi usa PocketTTS. Installa XTTS solo se ti serve.
echo ------------------------------------------------------------
set /p XTTS="Installare anche XTTS v2 (non commerciale)? [s/N]: "
if /i "!XTTS!"=="s" (
  echo Installo coqui-tts ...
  "%VPY%" -m pip install coqui-tts
  echo.
  echo NOTA: al primo utilizzo XTTS chiede di accettare la licenza CPML
  echo ^(non commerciale^). L'app la accetta in automatico solo perche' hai
  echo scelto tu questo motore per uso personale.
) else (
  echo Salto XTTS. Potrai installarlo dopo con:  pip install coqui-tts
)
echo.

REM --- 5. Controllo ffmpeg ------------------------------------
where ffmpeg >nul 2>&1
if errorlevel 1 (
  echo ------------------------------------------------------------
  echo   ffmpeg NON trovato.
  echo   Serve per le voci clonate e per montare il video.
  echo   Le voci Microsoft funzionano anche senza, ma il video no.
  echo ------------------------------------------------------------
  where winget >nul 2>&1
  if not errorlevel 1 (
    set /p FF="Installare ffmpeg adesso con winget? [S/n]: "
    if /i not "!FF!"=="n" winget install --id Gyan.FFmpeg -e --source winget
  ) else (
    echo Scarica ffmpeg da https://www.gyan.dev/ffmpeg/builds/
    echo poi aggiungi la cartella bin\ al PATH di Windows.
  )
) else (
  echo ffmpeg trovato.
)
echo.

REM --- 6. Controllo LibreOffice (serve per il VIDEO) ----------
set "HAS_LO="
where soffice >nul 2>&1 && set "HAS_LO=1"
if not defined HAS_LO if exist "C:\Program Files\LibreOffice\program\soffice.exe" set "HAS_LO=1"
if defined HAS_LO (
  echo LibreOffice trovato.
) else (
  echo ------------------------------------------------------------
  echo   LibreOffice NON trovato.
  echo   Serve SOLO per generare il VIDEO ^(disegna le slide^).
  echo   Per il solo PowerPoint con audio non e' necessario.
  echo ------------------------------------------------------------
  where winget >nul 2>&1
  if not errorlevel 1 (
    set /p LO="Installare LibreOffice adesso con winget? [S/n]: "
    if /i not "!LO!"=="n" winget install --id TheDocumentFoundation.LibreOffice -e --source winget
  ) else (
    echo Scarica LibreOffice da https://www.libreoffice.org/download/
  )
)
echo.

echo ============================================================
echo   Installazione completata.
echo   Avvia il programma con   Avvia_PPTX_TTS.bat
echo ============================================================
echo.
pause
