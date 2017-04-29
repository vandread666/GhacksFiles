Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, itemtype, mybox 

p = "HKCU\Control Panel\Mouse\SwapMouseButtons"
itemtype = "REG_SZ"
n = "1"
WSHShell.RegWrite p, n, itemtype

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", 64,"Done")
