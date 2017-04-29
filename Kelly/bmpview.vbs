Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, itemtype, mybox 

p = "HKCR\Paint.Picture\DefaultIcon\"
itemtype = "REG_SZ"
n = "%1"
WSHShell.RegWrite p, n, itemtype

MyBox = MsgBox("You must Reboot for the changes to take effect.", 48,"Done")
