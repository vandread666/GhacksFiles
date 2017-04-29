Dim WSHShell, n, MyBox, p, p1, t, cn, Caption, itemtype, errnum

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Help_Menu_Urls\Online_Support"
itemtype = "REG_SZ"
t = "Change Online Support Link"

On Error Resume Next

n = WSHShell.RegRead(p)

errnum = Err.Number

If errnum <> 0 then   

n = 0
End If
Caption = "Your current selection is: " & n & vbCR & vbCR
Caption = Caption & "Input the support link of your choice:"
On Error Goto 0

cn = InputBox(Caption, t, n)
If cn <> "" Then
WSHShell.RegWrite p, cn, itemtype
End If

If cn <>"" Then
MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")
End If