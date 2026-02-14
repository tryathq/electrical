@echo off
REM Builds the Windows installer (setup.exe).
REM 1. Run build_for_customer.bat first if you haven't (creates BackDownCalculator).
REM 2. Install Inno Setup from https://jrsoftware.org/isdl.php (free).
REM 3. Run this script. It will open the .iss in Inno Setup or compile if iscc is in PATH.

cd /d "%~dp0"

if not exist "BackDownCalculator" (
    echo BackDownCalculator folder not found.
    echo Please run build_for_customer.bat first, then run this again.
    pause
    exit /b 1
)

set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if not exist %ISCC% (
    echo Inno Setup 6 not found. Install from: https://jrsoftware.org/isdl.php
    echo Then either: run this script again, or open installer.iss in Inno Setup and click Build.
    start "" installer.iss
    pause
    exit /b 0
)

echo Building installer...
%ISCC% installer.iss
if %ERRORLEVEL% equ 0 (
    echo.
    echo Done. Installer: installer_output\BackDownCalculator_Setup.exe
    start "" "installer_output"
) else (
    echo Build failed.
)
pause
