@echo off
setlocal
cd /d "%~dp0"
title Installazione EgoMailExtractor

echo Installazione EgoMailExtractor...
echo Cartella: %CD%
echo.

rem Nel tuo ambiente usiamo prima "python". Non dipendiamo dal launcher "py".
where python >nul 2>nul
if errorlevel 1 goto :no_python

python --version
if errorlevel 1 goto :no_python

echo.
echo Verifica Tkinter...
python -c "import tkinter; print('Tkinter OK - versione Tcl/Tk', tkinter.TkVersion)"
if errorlevel 1 goto :no_tkinter

echo.
echo Creazione/aggiornamento ambiente virtuale .venv...
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto :error
) else (
    echo Ambiente .venv gia' presente: lo riutilizzo.
)

echo.
echo Aggiornamento pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error

echo.
echo Installazione dipendenze...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Verifica finale...
".venv\Scripts\python.exe" -c "import tkinter, cryptography, openpyxl, bs4; import egomail.app; print('Verifica dipendenze: OK')"
if errorlevel 1 goto :error

echo.
echo ============================================================
echo Installazione completata correttamente.
echo Ora avvia con Avvia_Windows.bat
echo ============================================================
echo.
pause
exit /b 0

:no_python
echo.
echo ERRORE: il comando python non e' disponibile.
echo Verifica l'installazione di Python e l'opzione Add Python to PATH.
echo.
pause
exit /b 1

:no_tkinter
echo.
echo ERRORE: Python funziona ma Tkinter non e' disponibile.
echo EgoMailExtractor necessita di Tkinter per l'interfaccia grafica.
echo Reinstalla Python includendo Tcl/Tk e IDLE.
echo.
pause
exit /b 1

:error
echo.
echo ============================================================
echo ERRORE durante l'installazione.
echo Leggi il messaggio visualizzato sopra.
echo ============================================================
echo.
pause
exit /b 1
