Set ShellApp = CreateObject("Shell.Application")
With ShellApp
	If MsgBox( .Namespace(&hd).Title & vbcrlf & .Namespace(&hd).self.path & vbcrlf & "( " & GetShName(.Namespace(&hd).self.path) & " )" & vbcrlf & vbcrlf & .Namespace(&he).Title & vbcrlf & .Namespace(&he).self.path & vbcrlf & "( " & GetShName(.Namespace(&he).self.path) & " )" & vbcrlf & vbcrlf & "Open these two folders in explorer now?", 68, "Check Desktop's Path") = 6 then
		.Open &hd
		.Open &he
	End If
End With

Function GetShName(folderspec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   GetShName = f.ShortPath
End Function