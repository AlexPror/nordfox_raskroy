@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [build] nordfox-raskroy — sdist + wheel в dist\
where python >nul 2>&1
if errorlevel 1 (
  echo Ошибка: не найден python в PATH.
  exit /b 1
)
python -m pip install -q "build>=1.0" wheel
if errorlevel 1 (
  echo Ошибка: не удалось установить пакеты build / wheel.
  exit /b 1
)
rem --no-isolation: не создаёт временный venv и не качает setuptools с PyPI
rem (нужен уже установленный setuptools в текущем Python).
python -m build --no-isolation
if errorlevel 1 (
  echo Ошибка: сборка завершилась с ошибкой.
  exit /b 1
)
echo Готово. Смотрите каталог dist\
dir /b dist\*.whl dist\*.tar.gz 2>nul
