Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
p = p & "IntelliMenus"
itemtype = "REG_SZ"
n = "No"

WSHShell.RegWrite p, n, itemtype
Title = "Personalized Menus are now disabled." & vbCR
Title = Title & "You may need to Log Off/Log On" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")