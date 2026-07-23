@echo off
setlocal

echo ============================================================
echo  BULK MAILER - MANUAL SETUP / REINSTALL
echo ============================================================
echo.
echo  You normally don't need this file - just double-click
echo  "Send Bulk Mail.bat" and it will set itself up automatically the
echo  first time. Use this file only if that didn't work, or if you
echo  want to force a clean reinstall of the required components.
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo.
    echo Please install Python first:
    echo   1. Go to https://www.python.org/downloads/
    echo   2. Download and run the installer.
    echo   3. IMPORTANT: On the first install screen, tick the box
    echo      that says "Add python.exe to PATH" before clicking Install.
    echo   4. Once installed, close this window and double-click
    echo      setup.bat again.
    echo.
    pause
    exit /b 1
)

echo Python found. Installing required components...
echo   - pywin32   (lets this tool talk to Outlook)
echo   - pandas    (reads your Excel file)
echo   - openpyxl  (Excel file support for pandas^)
echo.

python -m pip install --upgrade pip
python -m pip install pywin32 pandas openpyxl
if errorlevel 1 (
    echo.
    echo Something went wrong while installing. Common causes:
    echo   - No internet connection.
    echo   - Company network/proxy blocking pip. Ask IT for the proxy
    echo     settings, or ask IT to run this setup for you.
    echo.
    pause
    exit /b 1
)

echo.
echo Registering pywin32 with Windows (needed for Outlook automation)...
python -m pywin32_postinstall -install >nul 2>nul

echo.
echo Verifying installation...
python -c "import pandas, openpyxl, win32com.client, pythoncom; print('All components loaded successfully.')"
if errorlevel 1 (
    echo.
    echo Verification failed. Please re-run setup.bat, or share the
    echo error message above with your IT support.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete! You can now double-click "Send Bulk Mail.bat"
echo ============================================================
pause
