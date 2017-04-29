Dim WSHShell, RegKey, ScreenSaver, Result

Set WSHShell = CreateObject("WScript.Shell")

RegKey = "HKCU\Control Panel\Desktop\"

ScreenSaver = WSHShell.RegRead (regkey & "ScreenSaveActive")

If ScreenSaver = 1 Then 'Screen Saver is Enabled

Result = MsgBox("Your screen saver is currently active." & _ 
vbNewLine & "Would you like to disable it?", 36)

If Result = 6 Then 'clicked yes
WSHShell.RegWrite regkey & "ScreenSaveActive", 0
End If

Else 'Screen Saver is Disabled

Result = MsgBox("Your screen saver is currently disabled." & _ 
vbNewLine & "Would you like to enable it?", 36)

If Result = 6 Then 'clicked yes
WSHShell.RegWrite regkey & "ScreenSaveActive", 1
End If

End If
