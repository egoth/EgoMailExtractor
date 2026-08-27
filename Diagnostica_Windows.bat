@echo off
setlocal
cd /d "%~dp0"
title Diagnostica EgoMailExtractor

echo === Diagnostica EgoMailExtractor ===
echo Cartella: %CD%
echo.

echo [1] Python di sistema
where python
python --version

echo.
echo [2] Ambiente virtuale
if exist ".venv\Scripts\python.exe" (
    echo Trovato: .venv\Scripts\python.exe
    ".venv\Scripts\python.exe" --version
) else (
    echo NON TROVATO: .venv\Scripts\python.exe
)

echo.
echo [3] Tkinter e dipendenze
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import tkinter; print('Tkinter OK', tkinter.TkVersion); import cryptography, openpyxl, bs4; print('Dipendenze OK'); import egomail.app; print('EgoMailExtractor import OK')"
) else (
    python -c "import tkinter; print('Tkinter OK', tkinter.TkVersion)"
)

echo.
echo [4] File principali
if exist "main.py" (echo main.py OK) else (echo main.py MANCANTE)
if exist "requirements.txt" (echo requirements.txt OK) else (echo requirements.txt MANCANTE)

echo.
echo Fine diagnostica.
pause
