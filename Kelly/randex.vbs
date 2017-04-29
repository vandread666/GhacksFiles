'© Kelly Theriot and Doug Knox - 8/13/2003
'Kelly's Korner 
'http://www.kellys-korner-xp.com/xp_tweaks.htm

On Error Resume Next

Dim fname(3)
fname(0) = "\nstask32.exe"
fname(1) = "\winlogin.exe"
fname(2) = "\win32sockdrv.dll"
fname(3) = "\yuetyutr.dll"

Infected = 0

Set fso = CreateObject("Scripting.FileSystemObject")
Set tgrexe = fso.GetSpecialFolder(1)
Set tfolder = fso.GetSpecialFolder(2)
Set winfolder = fso.GetSpecialFolder(0)

For I = 0 to 3

If fso.FileExists(tgrexe & fname(I)) Then

   X = fso.DeleteFile(tgrexe & fname(I))
   Infected = Infected + 1

End If

Next

If Infected > 0 Then

X = fso.DeleteFolder(tfolder)

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\NDplDeamon"
WshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\MSConfig"
WshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\winlogon"
WshShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Runonce\NDplDeamon"
WshShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Runonce\NDpLDeamon"
WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Shell" , "explorer.exe"

End If

If Infected > 0 Then

Set MyFile = fso.OpenTextFile(winfolder & "\system.ini", 1)

Do While MyFile.AtEndOfStream = FALSE

LineIN = MyFile.ReadLine

If InStr(1,LineIn,"winlogin.exe",1) OR InStr(1,LineIn,"nstask32.exe",1) Then

Message = "You need to modify the Shell= line in SYSTEM.INI." & vbCR & vbCR
Message = Message & "It should read " & CHR(34) & "shell = explorer.exe" & CHR(34) & vbCR & vbCR
Message = Message & "Would you like to modify it now?"

SysIni = winfolder & "\SYSTEM.INI"

X = MsgBox(Message, vbYesNo, "Action required")

If X = 6 Then

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run("notepad " & SysIni)

End If

End If

Loop

MyFile.Close

End If

If Infected > 0 Then

Message = "W32.Randex.E Worm has been removed from your" & vbCr
Message = Message & "system. Update your antivirus definitions and install" & vbCR
Message = Message & "the Microsoft patch to prevent this in the future."

X = MsgBox(Message, vbOKOnly, "Done")

Else

Message = "The script did not find the W32.Randex E Worm file," & vbCR
Message = Message & "your system has not been cleaned."

X = MsgBox(Message, 4096, "No Repairs Made")

End If

Set WshShell = Nothing
Set fso = Nothing

On Error Goto 0
Y = MsgBox("Would you like to download the patch for Windows XP?", vbYesNo, "Important!")

If Y = 6 Then

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run("iexplore.exe http://www.microsoft.com/technet/treeview/default.asp?url=/technet/security/bulletin/MS03-039.asp")

End If

