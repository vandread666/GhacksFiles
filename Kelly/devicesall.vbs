Option Explicit

Dim WSHShell, n, p, itemtype

p = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\DEVMGR_SHOW_NONPRESENT_DEVICES"
itemtype = "REG_SZ"
n = "1"

Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite p, n, itemtype