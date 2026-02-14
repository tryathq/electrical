@echo off
REM Run this ONCE on your PC (Windows). It creates a folder you can zip and send to the customer.
REM Customer only: unzip and double-click START_APP.bat (no Python install needed on their side).

setlocal
cd /d "%~dp0"
set OUT=BackDownCalculator_Ready
set VENV=%OUT%\venv

echo Building customer package in: %OUT%
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"

echo Creating virtual environment...
python -m venv "%VENV%"
call "%VENV%\Scripts\activate.bat"
pip install -r requirements.txt -q

echo Copying app files...
copy app.py "%OUT%\"
copy config.py "%OUT%\"
copy reports_store.py "%OUT%\"
copy url_utils.py "%OUT%\"
copy instructions_parser.py "%OUT%\"
copy excel_builder.py "%OUT%\"
copy find_station_rows.py "%OUT%\"
copy requirements.txt "%OUT%\"
if exist CUSTOMER_README.txt copy CUSTOMER_README.txt "%OUT%\README.txt"
if exist .streamlit\config.toml (
    mkdir "%OUT%\.streamlit" 2>nul
    copy .streamlit\config.toml "%OUT%\.streamlit\"
)

mkdir "%OUT%\reports" 2>nul

echo Creating one-click launcher...
(
echo @echo off
echo cd /d "%%~dp0"
echo echo Starting app... browser will open shortly.
"%%~dp0venv\Scripts\streamlit.exe" run app.py
echo echo App is running. Close this window to stop the app.
echo pause
) > "%OUT%\START_APP.bat"

echo.
echo Done. Folder: %OUT%
echo.
echo Next: zip the folder "%OUT%" and send to customer.
echo Customer: unzip, then double-click START_APP.bat. No other steps.
pause
