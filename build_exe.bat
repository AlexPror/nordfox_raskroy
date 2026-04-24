@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"
echo [build_exe] Build desktop EXE (PyInstaller, onedir)

where py >nul 2>&1
if errorlevel 1 (
  echo ERROR: launcher "py" not found in PATH.
  exit /b 1
)

echo [1/5] Install/upgrade runtime dependencies...
py -3.11 -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo ERROR: failed to install requirements.
  exit /b 1
)

echo [2/5] Install/upgrade PyInstaller...
py -3.11 -m pip install -q "pyinstaller>=6.0"
if errorlevel 1 (
  echo ERROR: failed to install PyInstaller.
  exit /b 1
)

echo [3/5] Clean previous build artifacts...
if exist build rmdir /s /q build
if exist dist\nordfox_raskroy rmdir /s /q dist\nordfox_raskroy

echo [4/5] Build EXE...
py -3.11 -m PyInstaller ^
  --name nordfox_raskroy ^
  --noconsole ^
  --onedir ^
  --clean ^
  --paths src ^
  --collect-all PySide6 ^
  --collect-all reportlab ^
  --collect-all openpyxl ^
  src/nordfox_raskroy/__main__.py
if errorlevel 1 (
  echo ERROR: PyInstaller build failed.
  exit /b 1
)

echo [5/5] Verify output...
if exist "dist\nordfox_raskroy\nordfox_raskroy.exe" (
  echo SUCCESS: EXE built:
  echo   dist\nordfox_raskroy\nordfox_raskroy.exe
  echo.
  echo You can zip the entire folder:
  echo   dist\nordfox_raskroy\
  exit /b 0
)

echo ERROR: EXE not found in expected location.
exit /b 1
