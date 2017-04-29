On Error Resume Next

Dim WSHShell, n, MyBox, p, t, errnum, vers
Dim itemtype
Dim enab, disab, jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1803"

itemtype = "REG_DWORD"

enab = "ENABLED"
disab = "DISABLED"
jobfunc = "File Downloads are now "

t = "Confirmation"
Err.Clear
n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 0, itemtype
End If


If n = 0 Then
	n = 3
WSHShell.RegWrite p, n, itemtype
Mybox = MsgBox(jobfunc & disab & vbCR, 4096, t)
ElseIf n = 3 then
	n = 0
WSHShell.RegWrite p, n, itemtype
Mybox = MsgBox(jobfunc & enab & vbCR, 4096, t)
End If

Set WshShell = Nothing

MsgBox "Finished." & vbcr & vbcr , 4096, "Done"

