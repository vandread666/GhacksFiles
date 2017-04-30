'run_ie_reinstall.vbs - Runs the Internet Explorer Setup
'© Doug Knox - 4/10/2002
'This code may be freely distributed/modified
'Downloaded from www.dougknox.com

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("rundll32.exe setupapi,InstallHinfSection DefaultInstall 132 %windir%\Inf\ie.inf")
