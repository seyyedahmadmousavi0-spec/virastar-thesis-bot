Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "c:\Users\sadegh\Downloads\bale-thesis-bot\bale-thesis-bot"
WshShell.Run "python bot.py", 0, False

