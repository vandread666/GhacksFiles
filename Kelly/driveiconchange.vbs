'driveiconchange.vbs - Changes the default drive icon for your hard Drives in Explorer
'© Kelly Theriot - 07/11/03


On Error Resume Next

Dim WSHShell, p, strDrive, strIcon

strDrive = Inputbox("Enter the drive letter you wish to change the icon for:","Enter Drive Letter","X")

If strDrive = "" Then
	WScript.Quit
End If

strIcon = Inputbox("Enter the path to the icon wanted:","Enter Drive letter","")

If strIcon = "" Then
	WScript.Quit
End If

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\"& strDrive & "\DefaultIcon\"
WshShell.RegWrite p, strIcon

Set WshShell = Nothing

MsgBox "You must log off/log on, or end process on Explorer.exe for the changes to take effect.", 4096,"Finished"