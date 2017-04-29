'xp_kbKB890830 - Prevents Windows Update from prompting you to install
'Hotfix KB890830 , after you have already installed it.
'© Kelly Theriot - 4/18/2005

Set WshShell = CreateObject("WScript.Shell")
X = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix\KB890830\Fix Description")

If X = "Windows XP Hotfix - KB890830" Then
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\KB890830\","Needed to prevent Win Update from prompting for this install"
Set WshShell = Nothing

Message = "Windows Update should no longer prompt" & vbCR
Message = Message & "for Hotfix KB890830 to be installed." & vbCR & vbCR
Message = Message & "© 2005 - Kelly Theriot "

MsgBox Message, vbOkOnly, "Finished!"

Else

Message = "The Registry was checked and it doesn't appear" & vbCR
Message = Message & "that Hotfix KB890830 was ever installed." & vbCR 
Message = Message & "no changes were made to your system." & vbCR & vbCR
Message = Message & "© 2005 - Kelly Theriot "

MsgBox Message, vbInformation, "No changes made!"

End If

Set WshShell = Nothing
