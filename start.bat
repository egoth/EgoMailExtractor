@echo off
setlocal
cd /d "%~dp0"
title EgoMailExtractor

echo Avvio EgoMailExtractor...
echo Cartella: %CD%
echo.

rem Usa SEMPRE il python.exe della venv creata dall'installazione.
rem Non usiamo pythonw.exe: in caso di errore vogliamo vedere il messaggio.
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "%~dp0main.py"
    set "RC=%ERRORLEVEL%"
    if not "%RC%"=="0" goto :runtime_error
    exit /b 0
)

echo ATTENZIONE: ambiente virtuale .venv non trovato.
echo Provo il comando python installato nel sistema...
echo.

where python >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0main.py"
    set "RC=%ERRORLEVEL%"
    if not "%RC%"=="0" goto :runtime_error
    exit /b 0
)

echo ERRORE: Python non trovato e ambiente .venv assente.
echo Esegui prima Installa_Windows.bat.
echo.
pause
exit /b 1

:runtime_error
echo.
echo ============================================================
echo EgoMailExtractor si e' chiuso con un errore.
echo Codice errore: %RC%
echo Copia o fotografa il messaggio mostrato sopra e inviamelo.
echo ============================================================
echo.
pause
exit /b %RC%
