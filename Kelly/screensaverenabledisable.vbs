Dim WSHShell, n, MyBox, p, p1, t, cn, Caption, itemtype, errnum

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_USERS\.DEFAULT\Control Panel\Desktop\ScreenSaveActive"
itemtype = "REG_DWORD"
t= "Choose Accordingly"


On Error Resume Next

n = ws.RegRead(p)

errnum = Err.Number
If errnum <> 0 then  

n = 0
End If 

Caption = "1 = Screensaver Enable, 0 = Screensaver Disable"
On Error Goto 0

cn = InputBox(Caption, t, n)
If cn <> "" Then
WSHShell.RegWrite p, cn, itemtype
End If

If cn <>"" Then
'MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")
End If

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub