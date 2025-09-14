param([int]$Port=8501)
Set-Location -Path $PSScriptRoot
if (-not (Test-Path .\venv)) { try { py -m venv venv } catch { python -m venv venv } }
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
$inUse = (Get-NetTCPConnection -State Listen -LocalPort 8501 -ErrorAction SilentlyContinue)
if ($inUse) { $Port = 8502 }
Start-Process powershell -ArgumentList "-NoExit","-Command","streamlit run app.py --server.port $Port --server.headless=false"
Start-Process "http://localhost:$Port"
