@echo off
chcp 65001 >nul
echo در حال افزودن ربات به استارتاپ ویندوز...
set SCRIPT="%~dp0run_hidden.vbs"
set SHORTCUT="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\VirastarBot.lnk"

powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT%');$s.TargetPath='wscript.exe';$s.Arguments='%SCRIPT%';$s.WorkingDirectory='%~dp0';$s.Save()"

echo.
echo ✅ با موفقیت انجام شد!
echo اکنون با هر بار روشن شدن یا ورود به ویندوز، ربات به صورت خودکار و کاملاً مخفی در پس‌زمینه اجرا می‌شود.
pause

