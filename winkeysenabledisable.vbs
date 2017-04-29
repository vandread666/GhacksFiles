Option Explicit
On Error Resume Next

Dim WSHShell, n, p, itemtype, MyBox, vbdefaultbutton
Set WSHShell = WScript.CreateObject("WScript.Shell")

p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoWinKeys"
itemtype = "REG_DWORD"


n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 1, itemtype
End If

If n = 0 Then
   WshShell.RegWrite p, 1, itemtype
   MyBox = MsgBox("WinKeys are now DISABLED", 64, "Disable WinKeys")
End If

If n = 1 Then 
   WshShell.Regwrite p, 0, itemtype
   MyBox = MsgBox("WinKeys are now ENABLED", 64, "Enable WinKeys")
End If

Set WshShell = Nothing

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub