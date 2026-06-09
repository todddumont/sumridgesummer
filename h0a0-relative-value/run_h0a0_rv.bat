@echo off
set HISTORY_DIR=P:\jmorris\ICE H0A0 Historical Index Data
.\.venv\Scripts\python.exe .\h0a0_relative_value.py --history-dir "%HISTORY_DIR%"
pause
