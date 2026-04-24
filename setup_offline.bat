@echo off
setlocal

cd /d "%~dp0"
echo [1/5] Create venv...
py -3.11 -m venv .venv
if errorlevel 1 goto :err

echo [2/5] Upgrade pip...
.\.venv\Scripts\python -m pip install --upgrade pip
if errorlevel 1 goto :err

echo [3/5] Prepare wheelhouse...
if not exist wheelhouse mkdir wheelhouse

echo [4/5] Download requirements...
.\.venv\Scripts\python -m pip download -r requirements.txt -d wheelhouse
if errorlevel 1 goto :mirror

echo [5/5] Download build tools...
.\.venv\Scripts\python -m pip download build wheel setuptools -d wheelhouse
if errorlevel 1 goto :mirror

echo.
echo DONE: offline packages are in .\wheelhouse
goto :ok

:mirror
echo.
echo Primary PyPI failed. Trying mirror...
.\.venv\Scripts\python -m pip download -r requirements.txt -d wheelhouse -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
if errorlevel 1 goto :err
.\.venv\Scripts\python -m pip download build wheel setuptools -d wheelhouse -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
if errorlevel 1 goto :err

echo.
echo DONE (mirror): offline packages are in .\wheelhouse
goto :ok

:err
echo.
echo ERROR: failed to prepare offline packages.
exit /b 1

:ok
exit /b 0
