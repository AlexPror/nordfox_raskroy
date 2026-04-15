@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src"

where python >nul 2>&1
if %errorlevel%==0 (
  python -m nordfox_raskroy
) else (
  py -3 -m nordfox_raskroy
)
if errorlevel 1 pause
