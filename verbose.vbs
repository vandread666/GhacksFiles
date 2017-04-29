Option Explicit

Dim WSHShell, n, MyBox, p, itemtype

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system\"
p = p & "verbosestatus"
itemtype = "REG_DWORD"
n = 1

WSHShell.RegWrite p, n, itemtype