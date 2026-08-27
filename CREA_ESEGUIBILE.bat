@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Ambiente build assente. Avvio CREA_AMBIENTE_BUILD.bat...
  call CREA_AMBIENTE_BUILD.bat
  if errorlevel 1 exit /b 1
)
call PULISCI_BUILD.bat /silent
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean EgoMailExtractor_Community.spec
if errorlevel 1 goto :error
if not exist "dist\EgoMailExtractor_Community\EgoMailExtractor_Community.exe" goto :error
echo.
echo Build completata: dist\EgoMailExtractor_Community\EgoMailExtractor_Community.exe
echo Ora esegui CREA_DISTRIBUZIONE.bat per creare la cartella finale.
pause
exit /b 0
:error
echo.
echo ERRORE durante la compilazione.
pause
exit /b 1
