$ErrorActionPreference = "Stop"

$HistoryDir = "P:\jmorris\ICE H0A0 Historical Index Data"
$PythonExe = ".\.venv\Scripts\python.exe"

if (!(Test-Path $PythonExe)) {
    Write-Host "Virtual environment not found. Create it first with:" -ForegroundColor Yellow
    Write-Host "python -m venv .venv"
    exit 1
}

& $PythonExe .\h0a0_relative_value.py --history-dir $HistoryDir
