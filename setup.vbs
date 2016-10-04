Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """backuper.bat"" ""param.txt""",0
set WshShell = Nothing

