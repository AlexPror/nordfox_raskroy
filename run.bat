@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src"

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m nordfox_raskroy
) else (
  py -3.11 -m nordfox_raskroy
  if errorlevel 1 python -m nordfox_raskroy
)
if errorlevel 1 pause
