'xp_kb828026_fix.vbs - Prevents Windows Update from prompting you to install
'Hotfix KB828026, after you have already installed it.
'© Doug Knox and Kelly Theriot - 8/23/2003, 2/14/2004

Set WshShell = CreateObject("WScript.Shell")
X = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix\KB828026\Fix Description")

If X = "Windows XP Hotfix - KB828026" Then
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\KB828026\","Needed to prevent Win Update from prompting for this install"
Set WshShell = Nothing

Message = "Windows Update should no longer prompt" & vbCR
Message = Message & "for Hotfix KB828026 to be installed." & vbCR & vbCR
Message = Message & "© 2003 - Doug Knox and Kelly Theriot "

MsgBox Message, vbOkOnly, "Finished!"

Else

Message = "The Registry was checked and it doesn't appear" & vbCR
Message = Message & "that Hotfix KB828026 was ever installed." & vbCR 
Message = Message & "no changes were made to your system." & vbCR & vbCR
Message = Message & "© 2003 - Doug Knox and Kelly Theriot "

MsgBox Message, vbInformation, "No changes made!"

End If

Set WshShell = Nothing
