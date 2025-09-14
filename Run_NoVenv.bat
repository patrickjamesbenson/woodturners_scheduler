@echo on
setlocal enableextensions
pushd "%~dp0"
set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD ( where python >nul 2>nul && set "PY_CMD=python" )
if not defined PY_CMD (
  echo [!] Python not found.
  pause
  exit /b 1
)
%PY_CMD% -m pip install --user --upgrade pip
%PY_CMD% -m pip install --user -r requirements.txt
set "PORT=8502"
%PY_CMD% -m streamlit run app.py --server.port %PORT% --server.headless false
popd
endlocal
