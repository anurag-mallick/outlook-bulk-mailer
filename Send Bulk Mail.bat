@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title Bulk Mailer

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo  ============================================================
    echo   Python is required and was not found on this computer.
    echo  ============================================================
    echo.
    echo    1. Go to https://www.python.org/downloads/
    echo    2. Download and run the installer.
    echo    3. IMPORTANT: on the first install screen, tick the box that
    echo       says "Add python.exe to PATH" before clicking Install.
    echo    4. Once installed, close this window and double-click this
    echo       file again.
    echo.
    pause
    exit /b 1
)

python -c "import pandas, openpyxl, win32com.client, pythoncom" >nul 2>nul
if errorlevel 1 (
    echo.
    echo  ============================================================
    echo   First time here - setting things up. This only happens once
    echo   and takes a minute or two.
    echo  ============================================================
    echo.
    python -m pip install --upgrade pip >nul 2>nul
    python -m pip install pywin32 pandas openpyxl
    if errorlevel 1 (
        echo.
        echo  Something went wrong while installing. Common causes:
        echo    - No internet connection.
        echo    - Company network/proxy blocking pip. Ask your IT team for
        echo      help, or have them run this same file for you.
        echo.
        pause
        exit /b 1
    )
    python -m pywin32_postinstall -install >nul 2>nul

    python -c "import pandas, openpyxl, win32com.client, pythoncom" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo  Setup finished, but something still isn't right. Please
        echo  double-click this file once more - if it still fails, share
        echo  the message above with your IT support.
        echo.
        pause
        exit /b 1
    )

    echo.
    echo  Setup complete! Starting the Bulk Mailer now...
    echo.
)

where pythonw >nul 2>nul
if errorlevel 1 (
    start "" python bulk_mailer_gui.py
) else (
    start "" pythonw bulk_mailer_gui.py
)
