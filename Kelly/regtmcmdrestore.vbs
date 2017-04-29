Set WshShell = WScript.CreateObject("WScript.Shell")
With WScript.CreateObject("WScript.Shell")

On Error Resume Next

.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
.RegDelete "HKCU\Software\Policies\Microsoft\Windows\System\DisableCMD"
.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr"
.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system\DisableTaskMgr"

End With

Mybox = MsgBox(jobfunc & enab & vbCR & "Finished!", 4096, t)