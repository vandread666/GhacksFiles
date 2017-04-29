'Kelly's Korner 2003
'Special thanks to MVP Bill James 

Set oReg = GetObject("winmgmts:!root/default:StdRegProv")
Const HKLM = &H80000002
RegKeySUF = "SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder"
RegKeySUR = "SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg"

ResultsSUF = EnumKey(HKLM, RegKeySUF, False)
If ResultsSUF = "" Then
  ResultsSUF = space(5) & "(none)"
  iBtns = 0
Else
  iBtns = 4
  sDelPrompt = "Would you like to selectively delete any of these items?"
End If
sResults = "Disabled items in startupfolder key:" & vbcrlf & _
           ResultsSUF & vbcrlf & vbcrlf

ResultsSUR = EnumKey(HKLM, RegKeySUR, False)
If ResultsSUR = "" Then
  ResultsSUR = space(5) & "(none)"
  If iBtns <> 4 Then iBtns = 0
Else
  iBtns = 4
  sDelPrompt = "Would you like to selectively delete any of these items?"
End If
sResults = sResults & "Disabled items in startupreg key:" & vbcrlf & _
           ResultsSUR & vbcrlf & vbcrlf

If MsgBox(sResults & sDelPrompt, iBtns + 256) <> 6 Then WScript.quit

EnumKey HKLM, RegKeySUF, True
EnumKey HKLM, RegKeySUR, True

Function EnumKey(Key, SubKey, bDelete)
	Dim Ret()
	oReg.EnumKey Key, SubKey, sKeys

  On Error Resume Next
  
	ReDim Ret(UBound(sKeys))
  If Err = 13 Then Exit Function
  On Error GoTo 0

	For Count = 0 to UBound(sKeys)
	  If Not bDelete Then
	    'this branch used on first call
  		Ret(Count) = space(5) & sKeys(Count)
		Else
		  'this branch used on deletion iteration
		  If MsgBox("Do you want to delete " & sKeys(Count) & "?" & vbcrlf & _
		            vbcrlf & "This operation cannot be undone!", 4 + 256) = 6 Then
        DeleteKey HKLM, SubKey & "\" & sKeys(Count)
		  End If
  	End If
	Next
	EnumKey = Join(Ret, vbcrlf)
End Function

Function DeleteKey(Key, SubKey)
	DeleteKey = oReg.DeleteKey(Key, SubKey)
End Function

Set ws = WScript.CreateObject("WScript.Shell") 

On Error Resume Next
Err.Clear

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\MSConfig"

If Err.Number <> 0 then
MsgBox "You may need Administrator permissions to run this script" & vbcr & "or the registry entry does not exist.",4096,"Error!"
Else
MsgBox "The registry entry has been removed.", 4096,"Done!"
End If
Set WshShell = Nothing
