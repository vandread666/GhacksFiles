Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, itemtype, mybox

p = "HKEY_USERS\.DEFAULT\Control Panel\PowerCfg\CurrentPowerPolicy"
itemtype = "REG_SZ"
n = "4"
WSHShell.RegWrite p, n, itemtype

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", 64,"Done")