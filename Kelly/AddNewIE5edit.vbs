'AddNewIE5edit.vbs - Requires Internet Explorer 5+.
'Input for editor to add to "Edit with" button in IE. Also adds
'editor as right click option for htm/html files in Windows Explorer.
'Option to set editor as default for "View Source" in IE.
'© Bill James - bill@billsway.com
'rev 28/Oct/2001 18:11 - Fixed error if used with IE 6


Option Explicit
Dim WshShell,fso,Ed,Again,EdN,VSrc,Ttl,R1,R2,NL,Q
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Ttl = " - IE5 Edit Tool":NL = vbCRLF:Q = Chr(34)
R1 = "HKLM\Software\Microsoft\Internet Explorer\"
'
'Check IE5+.
If Not Left(WshShell.RegRead(R1 & "Version"),1) >= 5 Then
  MsgBox "Internet Explorer 5 or higher required",48,Ttl
  WScript.Quit
End If
'
'Get path to editor and validate.
Call GetInput
'
'Get editor name.
EdN = fso.GetBaseName(Ed)
R2 = "HKCR\htmlfile\shell\" & EdN & "\"
'
'Write values to Registry.
WshShell.RegWrite "HKCR\.htm\OpenWithList\" & EdN &_
  "\shell\edit\command\",Q & Ed & Q & " " & Q & "%1" & Q
WshShell.RegWrite R2,"Edit with &" & EdN
WshShell.RegWrite R2 & "command\",Q & Ed & Q & " " & Q & "%1" & Q
'
'Prompt to set as default for view source.
VSrc = MsgBox("Do you want to make " & EdN & " the default" &_
  NL & "for View Source in Internet Explorer?",4,Ttl)
If VSrc = 6 Then WshShell.RegWrite R1 & "View Source Editor\Editor Name\",Ed
'
MsgBox "Done.",64,Ttl
'
Sub GetInput
  Ed = InputBox("Enter path for editor to add for " &_
    "HTML files in IE5 and Windows Explorer.",Ttl)
  If Ed = "" Then WScript.Quit
  If Not fso.FileExists(Ed) Then
    Again = MsgBox("Invalid path or invalid characters used:  " &_
      Ed & NL & NL & "Try Again?",52,Ttl)
    If Again = 6 Then GetInput
  End If
End Sub
