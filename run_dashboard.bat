@echo off
setlocal
cd /d "%~dp0"
title UniAttend AI Dashboard
echo ========================================================
echo         Starting UniAttend AI Dashboard...
echo ========================================================
echo.
C:\Python310\python.exe run.py
if %errorlevel% neq 0 (
    echo.
    echo Application exited with an error.
    pause
)
