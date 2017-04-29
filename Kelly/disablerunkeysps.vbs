Option Explicit
On Error Resume Next

Dim WSHShell, n, p, p2, itemtype, MyBox
Set WSHShell = WScript.CreateObject("WScript.Shell")

p= "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisableLocalMachineRun"

p2= "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisableLocalMachineRunOnce"

itemtype = "REG_DWORD"

n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 1, itemtype
        WSHShell.RegWrite p2, 1, itemtype
End If

If n = 0 Then
   WshShell.RegWrite p, 1, itemtype
   MyBox = MsgBox("System Runkeys are now DISABLED", 64, "Disable Run Commands Specified in the Registry")
End If

If n = 1 Then 
   WshShell.Regwrite p, 0, itemtype
   WSHShell.RegWrite p2, 0, itemtype
   MyBox = MsgBox("System Runkeys are now ENABLED", 64, "Disabe Run Commands Specified in the Registry")
End If

Set WshShell = Nothing