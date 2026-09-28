@echo off
chcp 65001 > nul
cd /d "%~dp0"
python -m pip install --quiet pandas openpyxl xlrd
python merge_excel.py %*
pause
