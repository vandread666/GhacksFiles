Option Explicit
On Error Resume Next

Dim WSHShell, rcmd
Dim jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")

rcmd = "RunDll32 advpack.dll,LaunchINFSection C:\WINDOWS\inf\msnetmtg.inf,NetMtg.Remove"

WshShell.Run(rcmd)