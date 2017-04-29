Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, n, p, itemtype, MyBox, Title, caption

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\gk\"
itemtype = "REG_SZ"
n = "http://www.google.com/search?as_q=%s&as_sitesearch=support.microsoft.com" 

Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\gk\ "
itemtype = "REG_SZ"
n = "+"
 
Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\gk\%"
itemtype = "REG_SZ"
n = "%25" 

Ws.RegWrite p, n, itemtype

p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\gk\&"
itemtype = "REG_SZ"
n = "%26" 

Ws.RegWrite p, n, itemtype


p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchUrl\gk\+"
itemtype = "REG_SZ"
n = "%2B" 

Ws.RegWrite p, n, itemtype
Caption = "Google Search for KB articles from the Address Bar"
Title = "To use, type in: gk (followed by a space) then Q######."
MyBox = MsgBox(Title,4096,"Finished")