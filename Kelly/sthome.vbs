Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\"
p = p & "limitblankpassworduse"
itemtype = "REG_DWORD"
n = 0

WSHShell.RegWrite p, n, itemtype
Title = "Your Scheduled Tasks Can Now be Run without a Password." & vbCR
Title = Title & "You may need to log off/log on" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")