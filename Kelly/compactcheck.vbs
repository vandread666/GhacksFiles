'Set Compact Check Count to 0
'© Kelly's Korner - 4/05

Dim myShell, OE_ID, OE_ID2, OE_SA_Key, p, p1, t, p2, p3, n, itemtype

Set myShell = CreateObject("WScript.Shell")

On Error Resume Next
OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = myShell.RegRead(OE_ID)
OE_SA_Key = "HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\Compact Check Count"


p = "HKCU\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\"
p = p & "Compact Check Count"
itemtype = "REG_DWORD"
n = "0"

myshell. RegWrite p, n, itemtype

t = "Set Compress Check Count to 0"


p2 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoSaveSettings"
itemtype = "REG_DWORD"
n = "0"
myShell.RegWrite p2, n, itemtype

myshell.RegRead p1 & ""


p3 = "HKCU\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\Dont Show Dialogs\"
p3 = p3 & "Compact Do not Ask Again"
itemtype = "REG_DWORD"
n = "1"
myshell. RegWrite p3, n, itemtype

MsgBox "Compact Check Count has been reset to 0. Hit F5 to refresh", 4096,"Finished"


VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		myshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
