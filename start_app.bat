@echo off
cd /d "%~dp0"
if not exist venv (
  py -m venv venv 2>nul || python -m venv venv
)
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
start "" cmd /k "streamlit run app.py --server.port 8501 --server.headless=false"
start "" "http://localhost:8501"
