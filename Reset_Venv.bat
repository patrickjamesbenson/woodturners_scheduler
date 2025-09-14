@echo off
setlocal
pushd "%~dp0"
if exist "venv" rmdir /s /q "venv"
echo Done.
popd
endlocal
pause
