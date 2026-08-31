@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Ambiente non trovato. Eseguire prima install.bat
    pause
    exit /b 1
)

set "ICON=%CD%\assets\egomail_icon.ico"
if not exist "%ICON%" (
    echo Icona applicazione non trovata: %ICON%
    pause
    exit /b 1
)

.venv\Scripts\python.exe -m pip install pyinstaller
if errorlevel 1 goto :error

rem Elimina gli artefatti precedenti: in questo modo PyInstaller ricrea sempre
rem la risorsa ICON del file PE invece di lasciare residui di build vecchie.
if exist "build" rmdir /s /q "build"
if exist "dist\EgoMailExtractor" rmdir /s /q "dist\EgoMailExtractor"
if exist "EgoMailExtractor.spec" del /q "EgoMailExtractor.spec"

.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --name "EgoMailExtractor" --add-data "builtin_profiles;builtin_profiles" --add-data "help;help" --add-data "assets;assets" --add-data "LICENSE.txt;." --icon "%ICON%" main.py
if errorlevel 1 goto :error

if not exist "dist\EgoMailExtractor\EgoMailExtractor.exe" goto :error

echo.
echo Eseguibile creato in dist\EgoMailExtractor\EgoMailExtractor.exe
echo Icona incorporata da: %ICON%
echo.
echo NOTA: se Esplora file mostra ancora una vecchia icona dopo aver sostituito
echo un eseguibile omonimo, chiudi e riapri la cartella oppure riavvia Esplora file.
pause
exit /b 0

:error
echo.
echo ERRORE durante la creazione dell'eseguibile.
pause
exit /b 1
