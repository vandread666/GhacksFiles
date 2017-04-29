Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, n, p, itemtype, MyBox, Title, vbdefaultbutton

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\g\"
itemtype = "REG_SZ"
n = "http://www.google.com/search?q=%s" 

Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\g\ "
itemtype = "REG_SZ"
n = "+"
 
Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\g\%"
itemtype = "REG_SZ"
n = "%25" 

Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\g\&"
itemtype = "REG_SZ"
n = "%26" 

Ws.RegWrite p, n, itemtype


p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\g\+"
itemtype = "REG_SZ"
n = "%2B" 

Ws.RegWrite p, n, itemtype
Title = "To use, type in: g followed by a space then the search term."
MyBox = MsgBox(Title,4096,"Google Search from the Address Bar")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
