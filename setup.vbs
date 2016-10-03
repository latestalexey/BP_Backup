Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """K:\GIT_Repositories\BP_Backup\backuper.bat"" ""K:\GIT_Repositories\BP_Backup\param.txt""",0
set WshShell = Nothing

