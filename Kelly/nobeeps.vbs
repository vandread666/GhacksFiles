Option Explicit
Dim WSHShell, n, MyBox, p, itemtype, Title


Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Control Panel\Sound\"
p = p & "Beep"
itemtype = "REG_SZ"
n= "No"

WSHShell.RegWrite p, n, itemtype
Title = "System Beeps are now Disabled." & vbCR
Title = Title & "You may need to Log off/Log on" & vbCR
MyBox = MsgBox(Title,4096,"Finished")