'Serenity Macros 	http://www.angelfire.com/biz/serenitymacros 
'David Candy	davidc@sia.net.au

On Error Resume Next
Dim Sh
Set Sh = WScript.CreateObject("WScript.Shell")
ReportErrors "Creating Shell"
FileExt = InputBox("Enter the extension to look up?","Look Up Extensions","")
If left(fileext,1) = "." Then fileext=Mid(fileext,2)
sh.Run "http://shell.windows.com/fileassoc/0409/xml/redir.asp?Ext=" & FileExt


Sub ReportErrors(strModuleName)
	If err.number<>0 then Msgbox "Error occured in " & strModuleName & " module of " & err.number& " - " & err.description & " type" , vbCritical + vbOKOnly, "Something unexpected"
	Err.clear
End Sub

