' Lanzador de FCMENU sin ventana de consola visible.
' Usado por el acceso directo del Escritorio (ver manual, Capítulo 1).
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCarpeta = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strCarpeta
objShell.Run """" & strCarpeta & "\.venv\Scripts\pythonw.exe"" -m migration.ui.main_menu", 0, False
