@echo off
setlocal ENABLEDELAYEDEXECUTION
cd /d "%~dp0"
if not exist venv (
  py -m venv venv 2>nul || python -m venv venv
)
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
set PORT=8501
for /f "tokens=1,2,3,4,5*" %%a in ('netstat -aon ^| findstr LISTENING ^| findstr :8501') do set PORT=8502
start "Men's Shed Scheduler" cmd /k "streamlit run app.py --server.port %PORT% --server.headless=false"
timeout /t 3 >nul
start "" "http://localhost:%PORT%"
echo If the browser didn't open, copy: http://localhost:%PORT%
pause
