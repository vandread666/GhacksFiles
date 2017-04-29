'xp_taskbar_desktop_fixall.vbs - Repairs the Taskbar when minimized programs don't show.
'© Kelly Theriot and Doug Knox - 8/22/2003

Set WSHShell = WScript.CreateObject("WScript.Shell")

Message = "To work correctly, the script will close" & vbCR
Message = Message & "and restart the Windows Explorer shell." & vbCR
Message = Message & "This will not harm your system." & vbCR & vbCR
Message = Message & "Continue?"

X = MsgBox(Message, vbYesNo, "Notice")

If X = 6 Then 

On Error Resume Next

WshShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2\"
WshShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StreamMRU\"
WshShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Desktop\"

WshShell.RegDelete "HKCU\Software\Microsoft\Internet Explorer\Explorer Bars\{32683183-48a0-441b-a342-7c2a440a9478}\BarSize"

P1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"

WshShell.RegWrite p1 & "NoBandCustomize", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoMovingBands", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoCloseDragDropBands", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoSetTaskbar", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoToolbarsOnTaskbar", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoSaveSettings",0,"REG_DWORD"
WshShell.RegWrite p1 & "NoToolbarsOnTaskbar", 0, "REG_DWORD"
WshShell.RegWrite p1 & "NoSetTaskbar",0,"REG_DWORD"
WshShell.RegWrite p1 & "NoActiveDesktop",0,"REG_DWORD"
WshShell.RegWrite p1 & "ClassicShell",0,"REG_DWORD"

p1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"

WshShell.RegWrite p1 & "NoCloseDragDropBands", 0, "REG_DWORD"
WshShell.RegDelete p1 & "NoMovingBands"

p1 = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Shell"

WshShell.RegWrite p1, "explorer.exe", "REG_SZ"

p1 = "HKCU\Software\Microsoft\Internet Explorer\Explorer Bars\{32683183-48a0-441b-a342-7c2a440a9478}\"
WshShell.RegDelete p1 & "BarSize"
WshShell.RegWrite p1, "Media Band", "REG_SZ"

On Error Goto 0

For Each Process in GetObject("winmgmts:"). _
	ExecQuery ("select * from Win32_Process where name='explorer.exe'")
   Process.terminate(0)
Next

MsgBox "Finished." & vbcr & vbcr & "© Kelly Theriot and Doug Knox", 4096, "Done"

Else

MsgBox "No changes were made to your system." & vbcr & vbcr & "© Kelly Theriot and Doug Knox", 4096, "User Cancelled"

End If