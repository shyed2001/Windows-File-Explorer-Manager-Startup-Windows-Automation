@echo off
REM startup_launcher.bat
REM ---------------------
REM Launches the Windows File Explorer Manager automation script.
REM Place a shortcut to this file in the Windows Startup folder so it runs
REM automatically when Windows starts.
REM
REM Startup folder location:
REM   %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
REM
REM Usage:
REM   Double-click startup_launcher.bat  -OR-
REM   Run from the command prompt: startup_launcher.bat

SETLOCAL

REM Change to the directory that contains this batch file so relative paths work.
cd /d "%~dp0"

REM Check that Python is available.
python --version >NUL 2>&1
IF ERRORLEVEL 1 (
    echo [ERROR] Python was not found on PATH.
    echo         Please install Python 3.8+ and add it to your PATH.
    pause
    exit /b 1
)

REM Run the File Explorer Manager.
echo [INFO] Starting Windows File Explorer Manager...
python file_explorer_manager.py --config config.json

ENDLOCAL
