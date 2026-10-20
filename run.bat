@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Create the project virtual environment first; see README.md.
    exit /b 1
)
".venv\Scripts\python.exe" -B -X utf8 bot.py
exit /b %errorlevel%
