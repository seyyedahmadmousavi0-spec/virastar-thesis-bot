Option Explicit
Dim shell, files, root, python, script, quoted
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
root = files.GetParentFolderName(WScript.ScriptFullName)
python = files.BuildPath(root, ".venv\Scripts\pythonw.exe")
script = files.BuildPath(root, "bot.py")
If Not files.FileExists(python) Then
    MsgBox "Create the project virtual environment first; see README.md.", 48, "Virastar"
    WScript.Quit 1
End If
shell.CurrentDirectory = root
quoted = Chr(34)
shell.Run quoted & python & quoted & " -B -X utf8 " & quoted & script & quoted, 0, False

