@echo on
setlocal enableextensions
pushd "%~dp0"
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo Python not found. Install Python 3.11+ and retry.
  pause
  exit /b 1
)
if not exist "venv\Scripts\python.exe" (
  %PY% -m venv venv
)
call "venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
set "PORT=8502"
python -m streamlit run app.py --server.port %PORT% --server.headless false
popd
endlocal
