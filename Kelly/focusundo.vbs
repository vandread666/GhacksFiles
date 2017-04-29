Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Control Panel\desktop\"
p = p & "ForegroundLockTimeout"
itemtype = "REG_DWORD"
n = 0

WSHShell.RegWrite p, n, itemtype
Title = "Applications are not Prevented from Stealing Focus." & vbCR
Title = Title & "You may need to Log Off/Log On" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")