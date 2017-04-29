'ReportClassName.vbs
'Retrieves the class name for a file extension
'
'Serenity Macros 	http://www.angelfire.com/biz/serenitymacros 
'David Candy	davidc@sia.net.au
'
On Error Resume Next
strExplain="ReportClassName retrieves the class name and description for a file extension." & vbCRLF & vbCRLF & "In the Start - Settings - Folder Options - File Types dialog box registered file types are listed by description. This program allows a file extension to be typed in and will return the class name and description so it's properties can be edited in the File Type dialog box." & vbCRLF & vbCRLF & "If the File Class is set not to appear in the File Type dialog box the program will prompt to remove the restriction." & vbCRLF & vbCRLF
strTitle="Report Class Name"

Dim Sh
Set Sh = WScript.CreateObject("WScript.Shell")
ReportErrors "Creating Shell"
Sh.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\" & Wscript.ScriptName & "\", Chr(34) & Wscript.ScriptFullName & Chr(34) 
Sh.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\" & Left(Wscript.ScriptName, Len(Wscript.ScriptName)-3) & "exe" & "\", Chr(34) & Wscript.ScriptFullName & Chr(34) 
ReportErrors "Updating App Paths"


If MsgBox (strExplain  & "Continue?", vbYesNo + vbInformation, strTitle) = 6 then
	FileExtension=InputBox ("Enter the file extension", strTitle)
	If FileExtension<>"" then
		If Left(FileExtension,1)<>"." Then FileExtension = "." & FileExtension
		FileClass=Sh.RegRead("HKCR\" & FileExtension & "\")
		If Err.Number=-2147024894 then 
			Err.Clear
			MsgBox FileExtension & " file extension is not listed in the registry.",  vbInformation, strTitle
		Else
			If FileClass="" Then
				MsgBox FileExtension & " file extension is listed in the registry, but has no class name.",  vbInformation, strTitle
			Else
				FileClassDescription=Sh.RegRead("HKCR\" & FileClass & "\")
				If Err.Number=-2147024894 then 
					Err.Clear
					MsgBox FileExtension & " file extension is listed in the registry with a " & FileClass & " class, but the class doesn't exist.",  vbInformation, strTitle
				Else
					If FileClassDescription="" then FileClassdescription=""
					EditFlags=Sh.RegRead("HKCR\" & FileClass & "\EditFlags")
					Msg1="File Extension" & vbTab & FileExtension & vbcrlf & "File Class" & vbTab & vbTab & FileClass & vbcrlf & "File Description" & vbTab & FileClassDescription & vbcrlf & vbcrlf
					If  Err.Number =  -2147024894  then
						MsgBox Msg1 & " Files of " & FileExtension & " will be listed in the File Types dialog box as " & FileClassDescription & ".", vbOkOnly + vbInformation, strTitle
						Err.Clear
					ElseIf Err.Number = 0 Then
						If  VarType(EditFlags) = 8204 Then
							If EditFlags(0) And 1  then 
								If MsgBox (Msg1 & "File class is set not to appear in the File Types dialog box" & vbcrlf & vbcrlf & "Do you want this class to appear?", vbYesNo + vbQuestion, strTitle) = 6 then
									EditFlags(0)=EditFlags(0)-1
									DataToWrite=Convert4BytesToInteger(EditFlags(0) , EditFlags(1) , EditFlags(2) , EditFlags(3))
									Sh.RegWrite "HKCR\" & FileClass & "\EditFlags", DataToWrite, "REG_BINARY"
									MsgBox Msg1 & "Files of " & FileExtension & " will be listed in the File Types dialog box as " & FileClassDescription & ".", vbOkOnly + vbInformation, strTitle
								End If
							Else
								MsgBox Msg1 & "Files of " & FileExtension & " will be listed in the File Types dialog box as " & FileClassDescription & ".", vbOkOnly + vbInformation, strTitle
							End If
						ElseIf VarType(EditFlags)=3 then
							If EditFlags And 1  then 
								If MsgBox (Msg1 & "File class is set not to appear in the File Types dialog box" & vbcrlf & vbcrlf & "Do you want this class to appear?", vbYesNo + vbQuestion, strTitle) = 6 then
									EditFlags=EditFlags-1
									Sh.RegWrite "HKCR\" & FileClass & "\EditFlags", DataToWrite, "REG_DWORD"
									MsgBox Msg1 & "Files of " & FileExtension & " will be listed in the File Types dialog box as " & FileClassDescription & ".", vbOkOnly + vbInformation, strTitle
								End If
							Else
								MsgBox Msg1 & "Files of " & FileExtension & " will be listed in the File Types dialog box as " & FileClassDescription & ".", vbOkOnly + vbInformation, strTitle
							End If
						End If
					End If
				End If
			End If
		End If
	Else
		MsgBox "A blank value was entered or the dialog box was canceled.",  vbInformation, strTitle
	End If
End If
ReportErrors "General"

VisitSerenity

Sub ReportErrors(strModuleName)
	If err.number<>0 then Msgbox "An unexpected error occurred. This dialog provides details on the error." & vbCRLF & vbCRLF & "Error Details " & vbCRLF & vbCRLF & "Script Name" & vbTab & Wscript.ScriptFullName & vbCRLF & "Module" & vbtab & vbTab & strModuleName & vbCRLF & "Error Number" & vbTab & err.number & vbCRLF & "Description" & vbTab & err.description, vbCritical + vbOKOnly, "Something unexpected"
	Err.clear
End Sub

Sub VisitSerenity
	If MsgBox("This program came from the Serenity Macros Web Site" & vbCRLF & vbCRLF & "Would you like to visit Serenity's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton2, "Visit Serenity Macros") =6 Then
		sh.Run "http:\\www.angelfire.com\biz\serenitymacros"
	End If
End Sub

Function Convert4BytesToInteger(A,B,C,D) 
	If Len(D)=1 Then H="0" & D Else H=D
	If Len(C)=1 Then G="0" & C Else G=C
	If Len(B)=1 Then F="0" & B Else F=B
	If Len(A)=1 Then E="0" & A Else E=A
	Convert4BytesToInteger=CLng("&H" & H & G & F & E)
End Function

