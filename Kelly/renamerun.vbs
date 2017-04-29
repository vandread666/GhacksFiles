Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, Title 
Dim itemtype

p1 = "HKCR\CLSID\{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}\"

n = ws.RegRead(p1 & "LocalizedString")
t = "Change the Name of Run on Start Menu"
cn = InputBox("Type in the name to replace Run with.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "LocalizedString", cn
End If

Title = "Run on the Start Menu has now been renamed." 
MyBox = MsgBox(Title,64,"Finished")