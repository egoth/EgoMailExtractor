@echo off
setlocal
cd /d "%~dp0"

echo === EgoMailExtractor Community - preparazione ambiente build ===
echo Ricerca di Python 3.11, 3.12 o 3.13...

set "PYTHON_CMD="
set "PYTHON_ARGS="

rem 1) Python Launcher di Windows, se disponibile.
where py >nul 2>nul
if not errorlevel 1 (
  py -3.13 -c "import sys" >nul 2>nul && if not defined PYTHON_CMD set "PYTHON_CMD=py" && if not defined PYTHON_ARGS set "PYTHON_ARGS=-3.13"
  py -3.12 -c "import sys" >nul 2>nul && if not defined PYTHON_CMD set "PYTHON_CMD=py" && if not defined PYTHON_ARGS set "PYTHON_ARGS=-3.12"
  py -3.11 -c "import sys" >nul 2>nul && if not defined PYTHON_CMD set "PYTHON_CMD=py" && if not defined PYTHON_ARGS set "PYTHON_ARGS=-3.11"
)

rem 2) Installazione classica disponibile come "python".
if not defined PYTHON_CMD (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] <= (3,13) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
  )
)

rem 3) Alcune installazioni espongono il comando "python3" anche su Windows.
if not defined PYTHON_CMD (
  where python3 >nul 2>nul
  if not errorlevel 1 (
    python3 -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] <= (3,13) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python3"
  )
)

if not defined PYTHON_CMD goto :no_python

echo.
echo Python trovato:
%PYTHON_CMD% %PYTHON_ARGS% --version
echo Comando usato: %PYTHON_CMD% %PYTHON_ARGS%
echo.

if not exist ".venv\Scripts\python.exe" (
  echo Creazione ambiente virtuale .venv...
  %PYTHON_CMD% %PYTHON_ARGS% -m venv .venv
  if errorlevel 1 goto :error
) else (
  echo Ambiente virtuale .venv gia presente: verra riutilizzato.
)

.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto :error
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto :error
echo Verifica sintattica dei sorgenti...
.venv\Scripts\python.exe -m compileall -q egomail main.py
if errorlevel 1 goto :error

echo Esecuzione test automatici...
.venv\Scripts\python.exe -m pytest -q
if errorlevel 1 goto :error

echo.
echo Ambiente pronto e test superati.
pause
exit /b 0

:no_python
echo.
echo ERRORE: non e stata trovata un'installazione utilizzabile di Python 3.11, 3.12 o 3.13.
echo.
echo Ho provato automaticamente i comandi:
echo   py -3.13 / py -3.12 / py -3.11
echo   python
echo   python3
echo.
echo Prima di reinstallare Python, apri Prompt dei comandi ed esegui:
echo   python --version
echo   py --version
echo.
echo Se entrambi non funzionano, installa Python 3.12 64 bit da python.org.
echo Durante l'installazione seleziona "Add python.exe to PATH".
pause
exit /b 1

:error
echo.
echo ERRORE nella preparazione dell'ambiente.
pause
exit /b 1
