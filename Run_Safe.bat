@echo on
setlocal enableextensions
title Woodturners Scheduler - Safe Launcher

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
  if not exist "venv\Scripts\python.exe" (
    rmdir /s /q "venv"
  )
)
if not exist "venv\Scripts\python.exe" (
  echo [+] Creating virtual environment...
  %PY_CMD% -3.12 -m venv venv 2>nul || %PY_CMD% -3.11 -m venv venv 2>nul || %PY_CMD% -3 -m venv venv 2>nul || %PY_CMD% -m venv venv
  if errorlevel 1 (
     echo [!] Could not create venv. Falling back to no-venv mode...
     call "Run_NoVenv.bat"
     popd
     endlocal
     exit /b
  )
)
call "venv\Scripts\activate.bat"
python -m pip install --upgrade pip
if exist requirements.txt (
  pip install -r requirements.txt
) else (
  pip install streamlit==1.37.1 pandas==2.2.2 openpyxl==3.1.5 numpy python-dateutil pytz
)
set "PORT=8502"
python -m streamlit run app.py --server.port %PORT% --server.headless false
popd
endlocal
