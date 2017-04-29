Option Explicit

Dim  oShell, vData, vPath

Set oSHELL= WScript.CreateObject("WScript.Shell")
vData= ""
vPath= "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\"

On Error Resume Next
Do
    vData= oSHELL.RegRead(vPath & "LastKey"): If Err Then Exit Do
    If vData<>"My Computer" Then
        oSHELL.RegWrite vPath & "Favorites\Last Exit Key", vData, "REG_SZ"
        If Err Then Exit Do
        oSHELL.RegDelete vPath & "LastKey"
    End If
Loop Until True
Err.Clear: On Error GoTo 0

oSHELL.Run "regedit.exe", 1, False

WScript.Quit

