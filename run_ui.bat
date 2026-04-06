@echo off
cd /d "%~dp0"

set PYTHON=.venv\Scripts\python.exe

echo ========================================
echo STARTING STREAMLIT INVOICE APP
echo ========================================

%PYTHON% -m streamlit run streamlt_app.py

pause