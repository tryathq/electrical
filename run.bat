@echo off
cd /d "%~dp0"
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install -r requirements.txt -q
echo Opening app in browser...
streamlit run app.py
pause
