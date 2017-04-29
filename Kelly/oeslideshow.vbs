Dim myShell, OE_ID, OE_ID2, OE_SS_Key

Set myShell = CreateObject("WScript.Shell")

On Error Resume Next
OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = myShell.RegRead(OE_ID)

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, itemtype, mybox 

p = "HKCU\Identities\{F68D6F5C-259A-4C69-8819-724671E4B46F}\Software\Microsoft\Outlook Express\5.0\Autocatically Inline Images"
itemtype = "REG_DWORD"
n = "1"
WSHShell.RegWrite p, n, itemtype

p = "HKCU\Identities\{F68D6F5C-259A-4C69-8819-724671E4B46F}\Software\Microsoft\Outlook Express\5.0\Automatic Slide Show Delay"
itemtype = "REG_DWORD"
n = "2"
WSHShell.RegWrite p, n, itemtype

OE_SS_Key = "HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\"

myShell.RegWrite OE_SS_Key

MsgBox "Slide Show for OE in now enabled. You will need to restart OE for the changes to take effect.", 64,"Finished"


VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub


