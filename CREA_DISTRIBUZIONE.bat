@echo off
setlocal
cd /d "%~dp0"
if not exist "dist\EgoMailExtractor_Community\EgoMailExtractor_Community.exe" (
  echo Eseguibile non trovato. Avvio prima la compilazione...
  call CREA_ESEGUIBILE.bat
  if errorlevel 1 exit /b 1
)
if exist "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community" rmdir /s /q "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community"
mkdir "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community"
xcopy /e /i /y "dist\EgoMailExtractor_Community\*" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\" >nul
copy /y "LICENSE.txt" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\LICENSE.txt" >nul
copy /y "NOTICE" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\NOTICE" >nul
copy /y "LEGGIMI_DISTRIBUZIONE.txt" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\LEGGIMI.txt" >nul
copy /y "RELEASE_1.40.txt" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\RELEASE_1.40.txt" >nul
copy /y "RELEASE_1.40.2.txt" "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community\RELEASE_1.40.2.txt" >nul
echo.
echo Cartella pronta da distribuire:
echo %CD%\OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community
pause
