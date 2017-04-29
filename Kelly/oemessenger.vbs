Option Explicit

Dim WSHShell, MyBox, p1, q1 
Dim jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")
p1 = "HKEY_LOCAL_MACHINE\Software\Microsoft\Outlook Express\Hide Messenger"
q1 = 2

WSHShell.RegWrite p1, q1

