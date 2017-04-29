' Bill James, 30 Oct 2002

Set ws = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

RegTmp = ws.Environment("Process")("Temp") & "\RegTmp.tmp "
RegKey = "HKEY_CLASSES_ROOT\CLSID\{87D62D94-71B3-4b9a-9489-5FE6850DC73E}"

ws.Run "regedit /a /e " & RegTmp & RegKey, , True
Set RegFile = fso.OpenTextFile(RegTmp, 1)

NewRegTmp = ws.Environment("Process")("Temp") & "\NewRegTmp.tmp "
Set NewRegFile = fso.OpenTextFile(NewRegTmp, 2, True)

Do While Not RegFile.AtEndOfStream
  RegLine = RegFile.ReadLine
  If InStr(1, RegLine, RegKey, 1) <>  0 Then
    NewRegFile.WriteLine(Replace(RegLine, RegKey, RegKey & "-off"))
  Else
    
    NewRegFile.WriteLine(RegLine)
  End If
Loop

RegFile.Close
NewRegFile.Close

ws.Run "regedit /i /s " & NewRegTmp, , True

DelRegTmp = ws.Environment("Process")("Temp") & "\DelRegTmp.tmp "
Set DelRegFile = fso.OpenTextFile(DelRegTmp, 2, True)
DelRegFile.Write("REGEDIT4" & vbcrlf & "[-" & Mid(RegKey, 1) & "]")
DelRegFile.Close
ws.Run "regedit /i /s " & DelRegTmp, , True

fso.DeleteFile(RegTmp)
fso.DeleteFile(NewRegTmp)
fso.DeleteFile(DelRegTmp)
Set ws = Nothing
Set fso = Nothing
Set RegFile = Nothing
Set NewRegFile = Nothing
Set DelRegFile = Nothing

msgbox "Done"