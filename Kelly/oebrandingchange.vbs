Dim myShell, OE_ID, OE_ID2, OE_Tit_Key, p1, n, t, cn, mybox

Set ws = CreateObject("WScript.Shell")

On Error Resume Next

OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = ws.RegRead(OE_ID)
OE_Tit_Key = "HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\"

p1 = "HKCU\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\"

n = ws.RegRead(p1 & "WindowTitle")
t = "Change Name for OE"
cn = InputBox("Type in the name wanted.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "WindowTitle", cn
End If

MyBox = MsgBox("You must close and re-open OE to see the changes.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		ws.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub

