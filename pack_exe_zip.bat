@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

set "APP_DIR=dist\nordfox_raskroy"
set "APP_EXE=%APP_DIR%\nordfox_raskroy.exe"
set "VERSION_FILE=src\nordfox_raskroy\VERSION"

if not exist "%APP_EXE%" (
  echo ERROR: EXE not found: %APP_EXE%
  echo Run build_exe.bat first.
  exit /b 1
)

set "APP_VERSION=unknown"
if exist "%VERSION_FILE%" (
  for /f "usebackq delims=" %%v in ("%VERSION_FILE%") do (
    set "APP_VERSION=%%v"
    goto :version_done
  )
)
:version_done

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm"') do set "STAMP=%%i"
set "ZIP_NAME=nordfox_raskroy_%APP_VERSION%_desktop_%STAMP%.zip"
set "ZIP_PATH=dist\%ZIP_NAME%"

if exist "%ZIP_PATH%" del /q "%ZIP_PATH%"

echo Packing "%APP_DIR%" to "%ZIP_PATH%"...
powershell -NoProfile -Command "Compress-Archive -Path '%APP_DIR%\*' -DestinationPath '%ZIP_PATH%' -CompressionLevel Optimal"
if errorlevel 1 (
  echo ERROR: failed to create ZIP archive.
  exit /b 1
)

echo SUCCESS: archive created
echo   %ZIP_PATH%
exit /b 0
