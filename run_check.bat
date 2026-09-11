@echo off
chcp 65001 > nul
echo ===================================================
echo Running SanSan vs Excel Match Verification...
echo ===================================================
python check_match.py
echo.
pause
