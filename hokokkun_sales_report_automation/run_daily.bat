@echo off
cd /d "%~dp0"
"venv\Scripts\python.exe" "scripts\run_daily.py" >> "logs\task_scheduler_output.log" 2>&1
