'Copyright 2003 - Kelly Theriot
'http://www.kellys-korner-xp.com/xp_tweaks.htm

On Error Resume Next

Dim WshShell, bKey, cn, p1, p2, mybox, entryname
Set WshShell = WScript.CreateObject("WScript.Shell")

p1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures"

p2 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures"

n = "C:\Documents and Settings"

t = "Add a Folder to the Start Menu - Top Right"
cn = InputBox("Enter the drive and folder to be listed. Note:  It will pick up the name of the folder.  Also, this replaces the My Pictures Folder on the Start Menu. To control the display views use the My Pictures options under Advanced.", t, n)
If cn <> "" Then
  entryname = cn
    wshshell.RegWrite p1, cn
    wshshell.RegWrite p2, cn, "REG_EXPAND_SZ"


bKey = WshShell.RegRead(p1 & "")
bKey = WshShell.RegRead(p2 & "")

Else
    WScript.Quit
End If


MsgBox entryname & " has been added to the Start Menu. You must Log Off/Log On for the changes to take effect or End Process on Explorer.exe.",4096,"Finished!"

