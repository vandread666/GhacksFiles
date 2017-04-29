Set ShellApp = CreateObject("Shell.Application")
With ShellApp
	If MsgBox( .Namespace(&h0).Title & vbcrlf & .Namespace(&h0).self.path & vbcrlf & "( " & GetShName(.Namespace(&h0).self.path) & " )" & vbcrlf & vbcrlf & .Namespace(&h19).Title & vbcrlf & .Namespace(&h19).self.path & vbcrlf & "( " & GetShName(.Namespace(&h19).self.path) & " )" & vbcrlf & vbcrlf & "Open these two folders in explorer now?", 68, "Check Desktop's Path") = 6 then
		.Open 16
		.Open 25
	End If
End With

Function GetShName(folderspec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   GetShName = f.ShortPath
End Function