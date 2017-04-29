'Kelly's Korner 2004

Dim myShell, OE_ID, OE_ID2, OE_SA_Key, p1, n, itemtype, bkey

Set myShell = CreateObject("WScript.Shell")

On Error Resume Next
OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = myShell.RegRead(OE_ID)
OE_SA_Key = "HKEY_CURRENT_USER\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\Read in Plain Text only"


p1 = "HKCU\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\Read in Plain Text only"

itemtype = "REG_DWORD"
n = "1"

myShell.RegWrite p1, n, itemtype
MsgBox "Read in Plain Text only has been set.", 4096,"Finished"