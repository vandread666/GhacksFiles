Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\"
p = p & "FavIntelliMenus"
itemtype = "REG_SZ"
n = "Yes"

WSHShell.RegWrite p, n, itemtype
Title = "Personalized Menus are now enabled." & vbCR
Title = Title & "You may need to Log Off/Log On" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")