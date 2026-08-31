@echo off
setlocal
cd /d "%~dp0"
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community" rmdir /s /q "OUTPUT_DISTRIBUZIONE\EgoMailExtractor_Community"
if /i "%1"=="/silent" exit /b 0
echo Build temporanee eliminate.
pause
