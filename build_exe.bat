@echo off
REM Build .exe with PyInstaller, then zip the folder for the customer.
REM Customer: unzip and double-click BackDownCalculator.exe (no Python needed).

setlocal
cd /d "%~dp0"

echo Installing PyInstaller if needed...
pip install pyinstaller -q

echo.
echo Building .exe (this may take a few minutes)...
pyinstaller --noconfirm BackDownCalculator.spec
if errorlevel 1 (
    echo PyInstaller failed.
    pause
    exit /b 1
)

set OUT=dist\ElectricalReport
set ZIP=ElectricalReport_Ready.zip

echo.
echo Creating zip for customer: %ZIP%
if exist "%ZIP%" del "%ZIP%"
powershell -NoProfile -Command "Compress-Archive -Path '%OUT%' -DestinationPath '%ZIP%' -Force"
if errorlevel 1 (
    echo Zip failed. You can still send the folder: %OUT%
    pause
    exit /b 1
)

echo.
echo Done.
echo   Folder: %OUT%
echo   Zip:    %ZIP%
echo.
echo Send %ZIP% to the customer. They unzip and double-click BackDownCalculator.exe.
pause
