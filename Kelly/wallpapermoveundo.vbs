Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, MyBox 

Ws.RegDelete "HKCU\Control Panel\Desktop\Wallpaperoriginx"
Ws.RegDelete "HKCU\Control Panel\Desktop\Wallpaperoriginy" 

MyBox = MsgBox("Wallpaper modifications have been removed.", 64,"Done")