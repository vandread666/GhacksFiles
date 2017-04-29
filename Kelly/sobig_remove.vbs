'© Kelly Theriot and Doug Knox - 8/13/2003
'Special thanks to Bill James and Mike Kolitz
'Kelly's Korner 
'http://www.kellys-korner-xp.com/xp_tweaks.htm

Set fso = CreateObject("Scripting.FileSystemObject")
Set trgexe = fso.GetSpecialFolder(1)
bInfected = False

If (fso.FileExists(trgexe & "\winppr32.exe")) Then
  bInfected = True
End If

If (fso.FileExists(trgexe & "\winsst32.dat")) Then
  bInfected = True
End If

If NOT bInfected Then
  MsgBox "The script did not find the W32.Sobig.F Worm file, your system has not been cleaned.", 4096, "No Repairs Made"
Else
  KillWorms()
  MsgBox "The W32.Sobig.F Worm was removed from your system." & vbCR & "Update your antivirus definitions and install the Microsoft patch to prevent this in the future.", vbOKOnly, "Done"
End If

Set WshShell = Nothing
Set fso = Nothing

Public Function killWorms()
  For Each Process in GetObject("winmgmts:"). _
      ExecQuery ("select name from Win32_Process where name='winppr32.exe'")
    Process.terminate(0)
  Next

  On Error Resume Next

  fso.DeleteFile trgexe & "\winppr32.exe",True
  fso.DeleteFile trgexe & "\winsst32.dat",True

  With WScript.CreateObject("WScript.Shell")
    .RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\TrayX"
  End With

  On Error GoTo 0
End Function

