@echo off
chcp 65001 >nul
cd /d "%~dp0.."
echo Установка PySide6 (нужен интернет)...
python -m pip install "PySide6>=6.5"
pause
