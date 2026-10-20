@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
    echo Create the project virtual environment first; see README.md.
    exit /b 1
)
set "VIRASTAR_STARTUP_ROOT=%~dp0"
powershell -NoProfile -NonInteractive -Command "$root=$env:VIRASTAR_STARTUP_ROOT; $startup=[Environment]::GetFolderPath('Startup'); $link=Join-Path $startup 'VirastarBot.lnk'; $shell=New-Object -ComObject WScript.Shell; $shortcut=$shell.CreateShortcut($link); $shortcut.TargetPath=Join-Path $env:WINDIR 'System32\wscript.exe'; $shortcut.Arguments=([char]34)+(Join-Path $root 'run_hidden.vbs')+([char]34); $shortcut.WorkingDirectory=$root; $shortcut.Save()"
exit /b %errorlevel%
