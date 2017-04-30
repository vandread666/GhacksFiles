Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, vbdefaultbutton 

Ws.RegWrite "HKCU\Control Panel\Desktop\", "Wallpaperoriginx", "REG_SZ"
Ws.RegWrite "HKCU\Control Panel\Desktop\", "Wallpaperoriginy", "REG_SZ" 

p1 = "HKCU\Control Panel\Desktop\"

n = ws.RegRead(p1 & "")
t = "Change the Position of the Desktop Wallpaper"
cn = InputBox("Clear the box, then set the value of Wallpaperoriginx to equal the horizontal offset. Note: The image is offset from the center of your desktop, so you can use positive and negative values.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "Wallpaperoriginx", cn

End If

p1 = "HKCU\Control Panel\Desktop\"

n = ws.RegRead(p1 & "")
t = "Change the Position of the Desktop Wallpaper"
cn = InputBox("Clear the box, then set the value of Wallpaperoriginy to equal the vertical offset. Note: The image is offset from the center of your desktop, so you can use positive and negative values.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "Wallpaperoriginy", cn
End If

MyBox = MsgBox("Change your background image for the changes to take effect.", 64,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub

