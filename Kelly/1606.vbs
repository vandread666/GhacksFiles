Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title
Dim FSO, WinFolder

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WinFolder = fso.GetSpecialFolder(0)
Set FSO = Nothing

WinFolder = Left(WinFolder,3)


Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\"
p = p & "Common Administrative Tools"
itemtype = "REG_SZ"
n = WinFolder & "Documents and Settings\All Users\Start Menu\Programs\Administrative Tools"

WSHShell.RegWrite p, n, itemtype
Title = "Common Administrative Tools are now reset." & vbCR
Title = Title & "You may need to Log off/Log on or" & vbCR
Title = Title & "reboot for the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")