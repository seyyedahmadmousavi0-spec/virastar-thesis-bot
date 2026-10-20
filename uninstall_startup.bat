@echo off
setlocal
powershell -NoProfile -NonInteractive -Command "$link=Join-Path ([Environment]::GetFolderPath('Startup')) 'VirastarBot.lnk'; if (Test-Path -LiteralPath $link) { Remove-Item -LiteralPath $link -ErrorAction Stop }"
exit /b %errorlevel%
