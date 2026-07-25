@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

set "TAG_NAME=v2.9.0-tkinter"
set "DEV_BRANCH=feature/pyside6-ui"
set "TEMP_STATUS=%TEMP%\slidenarrator_git_status_%RANDOM%.txt"

cd /d "%~dp0"

echo.
echo ====================================================================
echo   SlideNarrator - Congelamento baseline Tkinter
echo ====================================================================
echo Tag stabile:       %TAG_NAME%
echo Branch nuova GUI:  %DEV_BRANCH%
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Git non e' disponibile nel PATH.
    goto :fail
)

if not exist ".git" (
    echo [ERRORE] Questa cartella non e' una repository Git.
    echo Esegui il file dalla cartella principale clonata da GitHub.
    goto :fail
)

for /f "delims=" %%B in ('git branch --show-current') do set "CURRENT_BRANCH=%%B"
if /I not "%CURRENT_BRANCH%"=="main" (
    echo [ERRORE] Il branch attivo e' "%CURRENT_BRANCH%", non "main".
    echo Torna su main e riprova: git switch main
    goto :fail
)

git status --porcelain > "%TEMP_STATUS%"
for %%A in ("%TEMP_STATUS%") do set "STATUS_SIZE=%%~zA"
del "%TEMP_STATUS%" >nul 2>&1
if not "%STATUS_SIZE%"=="0" (
    echo [ERRORE] La working tree contiene modifiche non registrate.
    echo Controlla prima con: git status
    echo Poi crea un commit completo della versione stabile.
    goto :fail
)

git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Il remote "origin" non e' configurato.
    goto :fail
)

echo [1/5] Aggiornamento riferimenti remoti...
git fetch origin --tags
if errorlevel 1 goto :git_fail

for /f "delims=" %%H in ('git rev-parse HEAD') do set "LOCAL_HEAD=%%H"
for /f "delims=" %%H in ('git rev-parse origin/main') do set "REMOTE_HEAD=%%H"
if /I not "%LOCAL_HEAD%"=="%REMOTE_HEAD%" (
    echo [ERRORE] main locale e origin/main non coincidono.
    echo Commit locale:  %LOCAL_HEAD%
    echo Commit remoto:  %REMOTE_HEAD%
    echo Esegui prima il push o risolvi la sincronizzazione.
    goto :fail
)

git show-ref --verify --quiet "refs/tags/%TAG_NAME%"
if not errorlevel 1 (
    echo [ERRORE] Il tag locale %TAG_NAME% esiste gia'.
    goto :fail
)

git ls-remote --exit-code --tags origin "refs/tags/%TAG_NAME%" >nul 2>&1
if not errorlevel 1 (
    echo [ERRORE] Il tag remoto %TAG_NAME% esiste gia'.
    goto :fail
)

git show-ref --verify --quiet "refs/heads/%DEV_BRANCH%"
if not errorlevel 1 (
    echo [ERRORE] Il branch locale %DEV_BRANCH% esiste gia'.
    goto :fail
)

git ls-remote --exit-code --heads origin "%DEV_BRANCH%" >nul 2>&1
if not errorlevel 1 (
    echo [ERRORE] Il branch remoto %DEV_BRANCH% esiste gia'.
    goto :fail
)

echo [2/5] Verifica manifest della baseline...
python "tools\verify_tkinter_baseline.py"
if errorlevel 1 goto :fail

echo [3/5] Creazione tag annotato...
git tag -a "%TAG_NAME%" -m "SlideNarrator 2.9.0 - baseline stabile Tkinter"
if errorlevel 1 goto :git_fail

echo [4/5] Pubblicazione tag...
git push origin "%TAG_NAME%"
if errorlevel 1 goto :git_fail

echo [5/5] Creazione e pubblicazione branch PySide6/QML...
git switch -c "%DEV_BRANCH%"
if errorlevel 1 goto :git_fail

git push -u origin "%DEV_BRANCH%"
if errorlevel 1 goto :git_fail

echo.
echo ====================================================================
echo   Baseline congelata correttamente
echo ====================================================================
echo Tag:     %TAG_NAME%
echo Branch:  %DEV_BRANCH%
echo Commit:  %LOCAL_HEAD%
echo.
echo Il prossimo lavoro deve avvenire nel branch %DEV_BRANCH%.
echo.
pause
exit /b 0

:git_fail
echo.
echo [ERRORE] Un comando Git non e' riuscito.
echo Controlla l'output precedente. Il codice sorgente non e' stato eliminato.

:fail
echo.
pause
exit /b 1
