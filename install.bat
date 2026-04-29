@echo off
setlocal enabledelayedexpansion
title Holiday Morning Messenger — Installer
cls

echo.
echo  =============================================
echo   Holiday Morning Messenger  ^|  Installer
echo  =============================================
echo.

:: ── 1. Find Python ──────────────────────────────────────────────
echo  [1/4]  Checking Python...

set PYTHON=
for %%P in (python python3) do (
    if "!PYTHON!"=="" (
        %%P --version >nul 2>&1 && set PYTHON=%%P
    )
)

if "!PYTHON!"=="" (
    echo.
    echo  ERROR: Python not found.
    echo.
    echo  Please install Python 3.8+ from https://www.python.org/downloads/
    echo  Make sure to tick "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

for /f "tokens=2" %%V in ('!PYTHON! --version 2^>^&1') do set PYVER=%%V
echo         Found Python !PYVER! — OK

:: ── 2. Check tkinter ─────────────────────────────────────────────
echo  [2/4]  Checking tkinter...
!PYTHON! -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ERROR: tkinter is not available.
    echo.
    echo  Re-install Python from python.org and make sure
    echo  "tcl/tk and IDLE" is checked in the Optional Features list.
    echo.
    pause
    exit /b 1
)
echo         tkinter — OK

:: ── 3. Desktop shortcut ──────────────────────────────────────────
echo  [3/4]  Creating desktop shortcut...
set APP_DIR=%~dp0
set APP_DIR=!APP_DIR:~0,-1!
set SHORTCUT=%USERPROFILE%\Desktop\Holiday Messenger.lnk

powershell -NoProfile -Command ^
 "$ws = New-Object -ComObject WScript.Shell;" ^
 "$sc = $ws.CreateShortcut('!SHORTCUT!');" ^
 "$sc.TargetPath = 'pythonw.exe';" ^
 "$sc.Arguments = '\"!APP_DIR!\holiday_messenger.py\"';" ^
 "$sc.WorkingDirectory = '!APP_DIR!';" ^
 "$sc.Description = 'Holiday Morning Messenger';" ^
 "$sc.Save()" >nul 2>&1

if exist "!SHORTCUT!" (
    echo         Desktop shortcut created — OK
) else (
    echo         Could not create shortcut ^(non-fatal^)
)

:: ── 4. Startup ───────────────────────────────────────────────────
echo  [4/4]  Setting up startup launch...
set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set STARTUP_FILE=!STARTUP_DIR!\holiday_messenger.bat

(
    echo @echo off
    echo start "" pythonw "!APP_DIR!\holiday_messenger.py"
) > "!STARTUP_FILE!"

if exist "!STARTUP_FILE!" (
    echo         Startup launcher installed — OK
) else (
    echo         Could not install startup launcher ^(non-fatal^)
)

:: ── Done ─────────────────────────────────────────────────────────
echo.
echo  =============================================
echo   Installation complete!
echo  =============================================
echo.
echo   Desktop shortcut :  Holiday Messenger
echo   Opens at login   :  Yes ^(Windows Startup folder^)
echo   App folder       :  !APP_DIR!
echo.
echo  Press any key to launch the app now...
pause >nul
start "" pythonw "!APP_DIR!\holiday_messenger.py"
