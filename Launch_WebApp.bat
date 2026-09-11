@echo off
chcp 65001 > nul
echo =========================================================
echo Launching SanSan Data Matcher Web Application...
echo =========================================================
echo Starting Streamlit server...
python -m streamlit run app.py
pause
