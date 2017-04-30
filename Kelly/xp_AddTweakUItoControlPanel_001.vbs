'© Doug Knox - 2003
'Based on AddTweakUItoCP.reg from http://www.kellys-korner-xp.com/xp_tweaks.htm

On Error Resume Next

Set WSHShell = WScript.CreateObject("WScript.Shell")

Set fso = CreateObject("Scripting.FileSystemObject")

SysDir = fso.GetSpecialFolder(1)
Tweak = SysDir & "\tweakui.exe"

If fso.FileExists(Tweak) Then

   Set WSHShell = WScript.CreateObject("WScript.Shell")

   p1 = "HKCR\CLSID\{D14ED2E1-C75B-443c-BD7C-FC03B2F08C17}\"

   WshShell.RegWrite p1, "TweakUIXP"
   WshShell.RegWrite p1 & "InfoTip", "Starts TweakUI for Windows XP"
   WshShell.RegWrite p1 & "DefaultIcon\","%SystemRoot%\\System32\\tweakui.exe,0", "REG_EXPAND_SZ"
   WshShell.RegWrite p1 & "Shell\Open\Command\","tweakui.exe"
   WshShell.RegWrite p1 & "ShellFolder\Attributes",48,"REG_DWORD"

   WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{D14ED2E1-C75B-443c-BD7C-FC03B2F08C17}\", "Add Tweakui to Control Panel"

MsgBox "TweakUI has been added to your Control Panel.", vbokonly, "Finished!"

Else

   MsgBox "TweakUI.EXE was not found in the " & SysDir & " folder. No changes were made.", vbokonly, "Problem!!"

End If

Set FSO = Nothing
Set WshShell = Nothing
