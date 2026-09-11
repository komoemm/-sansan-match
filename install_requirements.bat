@echo off
chcp 65001 > nul
echo =========================================================
echo Installing Required Python Packages for Web App...
echo လိုအပ်သော Python Library များကို ထည့်သွင်းနေပါသည်...
echo =========================================================
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
if %ERRORLEVEL% equ 0 (
    echo [အောင်မြင်ပါသည်] Package များ အားလုံး အောင်မြင်စွာ ထည့်သွင်းပြီးပါပြီ။
    echo ယခုအခါ 'Launch_WebApp.bat' ကို နှိပ်၍ အသုံးပြုနိုင်ပါပြီ။
) else (
    echo [သတိပေးချက်] Package ထည့်သွင်းရာတွင် အမှားအယွင်း ရှိနေပါသည်။
    echo ကျေးဇူးပြု၍ Python သွင်းထားခြင်း ရှိ/မရှိ နှင့် အင်တာနက် ချိတ်ဆက်မှု ရှိ/မရှိ စစ်ဆေးပေးပါ။
)
echo.
pause
