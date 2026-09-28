@echo off
chcp 65001 > nul
cd /d "%~dp0"
python -m pip install --quiet requests openpyxl
python recommend_keywords.py --limit 5 %*
pause
