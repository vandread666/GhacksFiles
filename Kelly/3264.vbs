'Copyright 2003 - Kelly Theriot and Doug Knox
'Modified by Mark Salloway - 07/09/2003
'http://www.kellys-korner-xp.com
'http://mvps.org/marksxp

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")

X = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")

If X = "x86" Then

Message = vbCR & "You have a 32 bit version of Windows" & vbCR & vbCR
Message = Message & "You should download 32 bit versions of Service Packs, " & vbCR
Message = Message & "Hotfixes and other Windows updates"

MsgBox Message, vbOkOnly, "32-Bit Windows Found"

Else

Message = vbCR & "Your Processor Architecture is not 32-Bit x86." & vbCR & vbCR
Message = Message & "This identifies a machine potentially running a 64-Bit" & vbCR 
Message = Message & "Edition of Windows. You should open your system properties" & vbCR
Message = Message & "to identify which hotfixes to apply on this system." & vbCR & vbCR
Message = Message & "Intel Itanium/II CPU = patches marked IA64" & vbCR
Message = Message & "AMD Opteron or Athlon64/FX = Patches marked as AMD64" & vbCR


MsgBox Message, vbOkOnly, "Non x86 Windows Found!"

End If
