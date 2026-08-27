@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" call CREA_AMBIENTE_BUILD.bat
if errorlevel 1 exit /b 1
.venv\Scripts\python.exe main.py
