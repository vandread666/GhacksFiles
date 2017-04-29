Option Explicit

Dim WSHShell, n, p, itemtype

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
p = p & "PowerdownAfterShutdown"
itemtype = "REG_SZ"
n = 0

WSHShell.RegWrite p, n, itemtype