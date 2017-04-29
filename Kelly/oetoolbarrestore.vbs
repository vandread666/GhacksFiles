'Copyright 2003 Kelly Theriot and Doug Knox

Option Explicit

Set myShell = CreateObject("WScript.Shell")
Dim myShell, OE_ID, OE_ID2, OE_Tit_Key, p, p1, n, itemtype

OE_ID = "HKEY_CURRENT_USER\Identities\Default User ID"
OE_ID2 = myShell.RegRead(OE_ID)

OE_Tit_Key = "HKCU\Identities\" & OE_ID2 & "\Software\Microsoft\Outlook Express\5.0\ShowToolbarIEAK"
myShell.RegWrite OE_Tit_Key, 1, "REG_DWORD"

MsgBox "The OE Toolbar has been Enabled.", 4096,"Finished"
