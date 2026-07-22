@echo off
cd /d "%~dp0"

where pythonw >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo Please double-click setup.bat first, then try again.
    pause
    exit /b 1
)

python -c "import pandas, openpyxl, win32com.client, pythoncom" >nul 2>nul
if errorlevel 1 (
    echo Required components are not installed yet.
    echo Please double-click setup.bat first, then try again.
    pause
    exit /b 1
)

start "" pythonw bulk_mailer_gui.py
