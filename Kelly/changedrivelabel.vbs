'changedrivelabel.vbs - Changes the default drive label for CD/DVD Drives in Explorer
'© Doug Knox and Kelly Theriot - 04/07/03

On Error Resume Next

'Declare variables
Dim WSHShell, p, strDrive, strLabel


strDrive = Inputbox("Enter the CD drive letter you wish to change the label for:","Enter Drive Letter","X")

If strDrive = "" Then
	WScript.Quit
End If

strLabel = Inputbox("Enter the label for this drive in Explorer","Enter Drive Label","CD Burner")

If strLabel = "" Then
	WScript.Quit
End If

'Set the Windows Script Host Shell and assign values to variables
Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\"& strDrive & "\DefaultLabel\"
WshShell.RegWrite p, strLabel

Set WshShell = Nothing

MsgBox "You must log off/log on, or reboot for the change to take effect.", 4096,"Finished"