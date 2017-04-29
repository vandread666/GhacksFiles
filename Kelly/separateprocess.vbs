Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
p = p & "DesktopProcess"
itemtype = "REG_DWORD"
n = 1

WSHShell.RegWrite p, n, itemtype
Title = "A Separate Process has been created." & vbCR
Title = Title & "You may need to log off/log on" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")