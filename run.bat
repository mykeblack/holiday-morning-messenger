@echo off
cd /d "%~dp0"

where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python was not found. Please install Python 3.10+ from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

python holiday_messenger.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo Myke's morning message closed with an error.
    echo.
    pause
)
