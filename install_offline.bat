@echo off
setlocal

cd /d "%~dp0"

if not exist wheelhouse (
  echo ERROR: .\wheelhouse not found. Run setup_offline.bat first.
  exit /b 1
)

echo [1/6] Recreate venv...
if exist .venv rmdir /s /q .venv
py -3.11 -m venv .venv
if errorlevel 1 goto :err

echo [2/6] Install runtime deps (offline)...
.\.venv\Scripts\python -m pip install --no-index --find-links wheelhouse -r requirements.txt
if errorlevel 1 goto :err

echo [3/6] Install build tools (offline)...
.\.venv\Scripts\python -m pip install --no-index --find-links wheelhouse build wheel
if errorlevel 1 goto :err

echo [4/6] Install project editable (offline)...
.\.venv\Scripts\python -m pip install --no-index --find-links wheelhouse -e .
if errorlevel 1 goto :err

echo [5/6] Quick import check...
.\.venv\Scripts\python -c "import openpyxl, reportlab, PySide6; print('imports_ok')"
if errorlevel 1 goto :err

echo [6/6] Done.
echo.
echo To run:
echo   .\.venv\Scripts\activate
echo   python -m nordfox_raskroy
exit /b 0

:err
echo.
echo ERROR: offline install failed.
exit /b 1
