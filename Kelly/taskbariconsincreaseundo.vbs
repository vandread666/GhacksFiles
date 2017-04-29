Option Explicit

Dim WSHShell, n, p, itemtype, MyBox

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\"
p = p & "MinWidth"
itemtype = "REG_SZ"
n = "-2310"

WSHShell.RegWrite p, n, itemtype

MyBox = MsgBox("You must reboot for the changes to take effect.", vbOKOnly,"Done")