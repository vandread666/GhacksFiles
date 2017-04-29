' PurgeCache V2.0 - 10/26/2001
' Copyright Walter Clayton, 2000 - all rights reserved
' This application may be distributed freely in unmodified form
' No warranties are expressed or implied.
' This script will extract the current location of the temporary internet files and
' set the system up to purge the contents of the folder on system reboot

Dim WSHSHell, fso, args, RSOreg, WinNTVer, Win9xVer, UserRunReg, NTOS, StartRegEntry, TIF

RSOreg = "HKLM\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce\DelTIF"
UserRunReg = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\DelTIF"
WinNTVer = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion"
Win9xVer = "HKLM\Software\Microsoft\Windows\CurrentVersion\VersionNumber"
NTOS = ""

Set WSHSHell = WScript.CreateObject ("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

On error resume next
NTOS = WSHShell.RegRead (WinNTVer)
If NTOS = "" then
  NTOS = False
  StartRegEntry = RSOreg
else
  NTOS = True
  StartRegEntry = UserRunReg
end if
On error goto 0

Set args = WScript.Arguments

TIF = ""
if args.count > 1 then
  if args (0) = "DeleteCache" then
    TIF = args (1)
  end if
end if

if TIF <> "" then
  DeleteCache (TIF)
else
  Configure
End if


Sub Configure

Dim TIF, TIFrg, RSO, Resp, L2, SN

TIFreg = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Cache"

TIF = WSHShell.RegRead (TIFreg)

Resp = MsgBox("The Cache located at: " & vbCR & vbCR & TIF & vbCR & vbCR & "Will be deleted on reboot." & vbCR & "OK to accept" & vbCR & "Cancel will cancel the request.", vbOKCancel)


If resp = vbCancel then
  On error resume next
  RSO = WSHShell.RegDelete (StartRegEntry)
  On error goto 0
else
  SN = Wscript.ScriptFullName
  L2 = "cscript """ & SN & """ //b //nologo DeleteCache " & """" & TIF & """"
  RSO = WSHShell.RegWrite (StartRegEntry, L2)
end if

End Sub


Sub DeleteCache (TIF)

Dim f1

on error resume next
fso.deletefolder (TIF) 
If NTOS then 
  RSO = WSHSHell.RegDelete (StartRegEntry) 
End if
on error goto 0

End Sub
