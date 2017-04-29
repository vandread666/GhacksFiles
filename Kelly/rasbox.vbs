Option Explicit

Dim WSHShell, MyBox, p1, q1 
Dim jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")
p1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\RasDisable"
q1 = 1

WSHShell.RegWrite p1, q1