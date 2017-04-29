Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, itemtype, mybox 

p = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Max Cached Icons"
itemtype = "REG_SZ"
n = "12000"
WSHShell.RegWrite p, n, itemtype

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", 48,"Done")
