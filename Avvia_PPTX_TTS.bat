@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [ERRORE] Ambiente non installato.
  echo Esegui prima  Installa_PPTX_TTS.bat  (un solo clic^).
  echo.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" pptx_tts_gui.py

if errorlevel 1 (
  echo.
  echo Si e' verificato un errore durante l'avvio.
  echo Controlla che tutti i file .py siano nella stessa cartella di questo
  echo file e di aver eseguito Installa_PPTX_TTS.bat.
  echo.
  pause
)
