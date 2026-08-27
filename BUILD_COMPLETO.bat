@echo off
setlocal
cd /d "%~dp0"
call CREA_AMBIENTE_BUILD.bat
if errorlevel 1 exit /b 1
call CREA_ESEGUIBILE.bat
if errorlevel 1 exit /b 1
call CREA_DISTRIBUZIONE.bat
