Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """E:\BP_Backup\backuper.bat"" ""E:\BP_Backup\param.txt""",0
set WshShell = Nothing

