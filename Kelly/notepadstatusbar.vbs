Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Notepad\"
p = p & "StatusBar"
itemtype = "REG_DWORD"
n = "1"

WSHShell.RegWrite p, n, itemtype
Title = "The Status Bar in now enabled w/o eliminating Word Wrap." & vbCR
Title = Title & "You may need to Log off/Log on. For the change to take effect." & vbCR
MyBox = MsgBox(Title,4096,"Finished")