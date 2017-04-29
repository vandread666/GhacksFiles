Message = "To work correctly, the script will close" & vbCR
Message = Message & "and restart the Windows Explorer shell." & vbCR
Message = Message & "This will not harm your system." & vbCR & vbCR
Message = Message & "Continue?"

X = MsgBox(Message, vbYesNo, "Notice")

If X = 6 Then

On Error Resume Next

Dim WSHShell, n, MyBox, p, t, errnum, vers
Dim itemtype
Dim enab, disab, jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarGlomming"

itemtype = "REG_DWORD"

enab = "ENABLED"
disab = "DISABLED"
jobfunc = "Taskbar Grouping is now "

t = "Confirmation"
Err.Clear
n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 1, itemtype
End If


If n = 1 Then
	n = 0
WSHShell.RegWrite p, n, itemtype
Mybox = MsgBox(jobfunc & disab & vbCR, 4096, t)
ElseIf n = 0 then
	n = 1
WSHShell.RegWrite p, n, itemtype
Mybox = MsgBox(jobfunc & enab & vbCR, 4096, t)
End If

Set WshShell = Nothing

On Error GoTo 0

For Each Process in GetObject("winmgmts:"). _
	ExecQuery ("select * from Win32_Process where name='explorer.exe'")
   Process.terminate(0)
Next

MsgBox "Finished." & vbcr & vbcr , 4096, "Done"

Else 

MsgBox "No changes were made to your system." & vbcr & vbcr, 4096, "User Cancelled"

End If