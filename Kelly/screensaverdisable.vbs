Option Explicit
Dim WSHShell, n, MyBox, p, itemtype, Title


Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_USERS\.DEFAULT\Control Panel\Desktop\"
p = p & "ScreenSaveActive"
itemtype = "REG_SZ"
n= 0

WSHShell.RegWrite p, n, itemtype
Title = "Screen Saver is now Disabled." & vbCR
Title = Title & "You may need to log off/log on" & vbCR
MyBox = MsgBox(Title,4096,"Finished")