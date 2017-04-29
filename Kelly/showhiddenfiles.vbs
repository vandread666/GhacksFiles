Message = "To work correctly, the script will close" & vbCR
Message = Message & "and restart the Windows Explorer shell." & vbCR
Message = Message & "This will not harm your system." & vbCR & vbCR
Message = Message & "Continue?"

X = MsgBox(Message, vbYesNo, "Notice")

If X = 6 Then
On Error Resume Next


On Error Resume Next

Dim WSHShell, n, p, itemtype, MyBox
Set WSHShell = WScript.CreateObject("WScript.Shell")

p = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"

itemtype = "REG_DWORD"

n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 1 then

	WSHShell.RegWrite p, 2, itemtype
End If

If n = 2 Then
   WshShell.RegWrite p, 1, itemtype
   MyBox = MsgBox("Show Hidden Files and Folders are now ENABLED", 64, "Hidden Files and Folders")
End If

If n = 1 Then 
   WshShell.Regwrite p, 2, itemtype
   MyBox = MsgBox("Show Hidden Files and Folders are now DISABLED", 64, "Hidden Files and Folders")
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