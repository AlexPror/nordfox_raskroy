@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src"
set "PY_EXE="

if exist ".venv\Scripts\python.exe" (
  set "PY_EXE=.venv\Scripts\python.exe"
)

if not defined PY_EXE (
  where py >nul 2>&1
  if not errorlevel 1 set "PY_EXE=py -3.11"
)

if not defined PY_EXE (
  where python >nul 2>&1
  if not errorlevel 1 set "PY_EXE=python"
)

if not defined PY_EXE (
  echo ERROR: Python not found. Install Python 3.11+ or create .venv.
  pause
  exit /b 1
)

%PY_EXE% -m nordfox_raskroy
if errorlevel 1 pause
