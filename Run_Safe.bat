@echo on
setlocal enableextensions
title WOTH Scheduler - Safe Launcher
pushd "%~dp0"
set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD ( where python >nul 2>nul && set "PY_CMD=python" )
if not defined PY_CMD (
  echo [!] Python was not found. Install Python 3.11+ or 3.12 and re-run.
  pause
  exit /b 1
)
if exist "venv" (
  if not exist "venv\Scripts\python.exe" rmdir /s /q "venv"
)
if not exist "venv\Scripts\python.exe" (
  echo [+] Creating virtual environment...
  %PY_CMD% -3.12 -m venv venv 2>nul || %PY_CMD% -3.11 -m venv venv 2>nul || %PY_CMD% -3 -m venv venv 2>nul || %PY_CMD% -m venv venv
)
call "venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
set "PORT=8502"
python -m streamlit run app.py --server.port %PORT% --server.headless false
popd
endlocal
