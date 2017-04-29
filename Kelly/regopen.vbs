'Regedit Opener - bypasses Regedit opening to the last viewed key
'regopen.vbs 
'This code may be freely distributed/modified

Option Explicit
On Error Resume Next

Dim WSHShell 
Set WSHShell=Wscript.CreateObject("Wscript.Shell") 
WSHShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\LastKey" 
WSHShell.Run "REGEDIT"

Set WSHShell = Nothing