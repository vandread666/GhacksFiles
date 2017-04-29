'Enable Outlook Express Splash Screen
'© Kelly's Korner - 7/29/03

Dim myShell, OE_ID, OE_ID2, OE_SP_Key, p1, cn, n, t, itemtype 

Set myShell = CreateObject("WScript.Shell")

On Error Resume Next
OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = myShell.RegRead(OE_ID)
OE_SP_Key = "HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\NoSplash"

p1="HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\NoSplash"

itemtype = "REG_DWORD"
n = "0"
myShell.RegWrite p1, n, itemtype


MsgBox "OE Splash Screen has been enabled.", 4096,"Finished"


VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		myshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub