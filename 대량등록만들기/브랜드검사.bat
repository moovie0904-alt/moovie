@echo off
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python -m pip install --quiet openpyxl
python brand_check.py %*
pause
