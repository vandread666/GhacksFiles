Option Explicit

Dim WSHShell, RegKey, ScreenSaver, Result

Set WSHShell = CreateObject("WScript.Shell")

RegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"

ScreenSaver = WSHShell.RegRead (regkey & "DisableCAD")

If ScreenSaver = 0 Then 'Ctrl/Alt/Delete is Enabled.

   Result = MsgBox("Ctrl/Alt/Delete is currently active." & _ 
        vbNewLine & "Would you like to disable it?", 36)

   If Result = 6 Then 'clicked yes
      WSHShell.RegWrite regkey & "DisableCAD", 1
   End If

Else 'Ctrl/Alt/Delete is Disabled

   Result = MsgBox("Ctrl/Alt/Delete is currently disabled." & _ 
        vbNewLine & "Would you like to enable it?", 36)

   If Result = 6 Then 'clicked yes
      WSHShell.RegWrite regkey & "DisableCAD", 0
   End If

End If
