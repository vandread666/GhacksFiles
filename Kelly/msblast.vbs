'© Kelly Theriot and Doug Knox - 8/13/2003
'Special thanks to Bill James and Mike Kolitz
'Kelly's Korner 
'http://www.kellys-korner-xp.com/xp_tweaks.htm

Set fso = CreateObject("Scripting.FileSystemObject")
Set trgexe = fso.GetSpecialFolder(1)
bInfected = False

If (fso.FileExists(trgexe & "\msblast.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\penis32.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\wuaumgr.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\teekids.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\root32.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\msconfig35.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\mspatch.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\mslaugh.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\enbiei.exe")) Then
   bInfected = True
End If

If NOT bInfected Then
  MsgBox "The script did not find the W32.Blaster.Worm file, your system has not been cleaned." & vbCR & vbCR & "Copyright 2003 - Kelly Theriot and Doug Knox",4096, "No Repairs Made"
Else
  KillWorms()
  MsgBox "One or more variants of W32.Blaster.Worm were removed from your system." & vbCR & "Update your antivirus definitions and install the Microsoft patch to prevent this in the future." & vbCR & vbCR & "Copyright 2003 - Kelly Theriot and Doug Knox", vbOKOnly, "Done"
End If

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")

If WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\KB828741\DisplayName") <> "Windows XP Hotfix - KB828741" Then

  X = MsgBox("Would you like to download the patch for Windows XP?", vbYesNo, "Important!")
  
  If X = 6 Then
          WshShell.Run("iexplore.exe http://www.microsoft.com/technet/security/bulletin/ms04-012.mspx")
  End If

Else

  MsgBox "You already have the patch for the Blaster Worm exploit installed.", vbOkOnly, "Finished"

End If

On Error Goto 0

Set WshShell = Nothing
Set fso = Nothing

Public Function killWorms()
  For Each Process in GetObject("winmgmts:"). _
      ExecQuery ("select name from Win32_Process where name='msblast.exe' OR name='penis32.exe' OR name='teekids.exe' OR name='wuaumgr.exe' OR name='root32.exe' OR name='msconfig35.exe' OR name='mspatch.exe' OR name='mslaugh.exe' OR name='enbiei.exe'")
    Process.terminate(0)
  Next

  On Error Resume Next

  fso.DeleteFile trgexe & "\msblast.exe",True
  fso.DeleteFile trgexe & "\penis32.exe",True
  fso.DeleteFile trgexe & "\teekids.exe",True
  fso.DeleteFile trgexe & "\wuaumgr.exe",True
  fso.DeleteFile trgexe & "\root32.exe",True
  fso.DeleteFile trgexe & "\msconfig35.exe",True
  fso.DeleteFile trgexe & "\mspatch.exe",True
  fso.DeleteFile trgexe & "\mslaugh.exe",True
  fso.DeleteFile trgexe & "\enbiei.exe",True

  With WScript.CreateObject("WScript.Shell")

    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\windows automation"
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\windows auto update"
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Microsoft Inet Xp.."
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\MSConfig"
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Nonton Antivirus"
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce\Nonton Antivirus"
    .RegDelete "HLKM\Software\Microsoft\Windows\CurrentVersion\Run\www.hidro.4t.com"

  End With

  On Error GoTo 0
End Function

