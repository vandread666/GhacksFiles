Option Explicit

Dim WSHShell, n, p, t, cn, vbdefaultbutton
Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32\Placesbar\"
t = "Add a Folder to the Places Bar"

Dim place, pl, c
c = 1

For Each place in Array("Place0","Place1","Place2","Place3","Place4")
  pl = p & place

On Error Resume Next
  n = WSHShell.RegRead(pl)
  If Err Then n = ""

 On Error GoTo 0
  cn = InputBox("Type the Path and click OK for Entry " & c, t, n)
  If cn <> "" Then WSHShell.RegWrite pl, cn

 c = c + 1
Next

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
