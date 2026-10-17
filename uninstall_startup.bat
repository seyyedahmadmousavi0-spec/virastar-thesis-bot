@echo off
chcp 65001 >nul
set SHORTCUT="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\VirastarBot.lnk"
if exist %SHORTCUT% (
    del %SHORTCUT%
    echo ✅ ربات از استارتاپ ویندوز حذف شد.
) else (
    echo ربات در استارتاپ وجود نداشت.
)
pause

