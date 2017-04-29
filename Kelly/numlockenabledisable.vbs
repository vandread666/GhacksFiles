Option Explicit
On Error Resume Next

Dim WSHShell, n, p, itemtype, MyBox, vbdefaultbutton
Set WSHShell = WScript.CreateObject("WScript.Shell")

p = "HKEY_CURRENT_USER\Control Panel\Keyboard\InitialKeyboardIndicators"
itemtype = "REG_SZ"

n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 2, itemtype
End If

If n = 0 Then
   WshShell.RegWrite p, 2, itemtype
   MyBox = MsgBox("Numlock is now ENABLED", 64, "Enable Numlock")
End If

If n = 2 Then 
   WshShell.Regwrite p, 0, itemtype
   MyBox = MsgBox("Numlock is now DISABLED", 64, "Enable Numlock")
End If

Set WshShell = Nothing

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub