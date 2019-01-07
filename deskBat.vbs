Dim fso,f,s
Set fso = CreateObject("Scripting.FileSystemObject")
flsr = "C:\oashurtico"
If fso.FolderExists(flsr) Then
s="exist"
Else
s="not exist"
set f = fso.CreateFolder(flsr)
End If
test = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
aa = test + "\oa.ico"
Set fs=WScript.CreateObject("scripting.filesystemobject")
fs.CopyFile aa, "C:\oashurtico\oa.ico"
Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
lnkname = strDesktop + "\OA.lnk"
set oShellLink = WshShell.CreateShortcut(lnkname)
oShellLink.TargetPath = "http://www.baidu.com"
oShellLink.WindowStyle = 3
oShellLink.IconLocation = "C:\oashurtico\oa.ico"
oShellLink.Description = "¿ì½Ý·½Ê½"
oShellLink.WorkingDirectory = "C:\oashurtico\oa.ico"
oShellLink.Save