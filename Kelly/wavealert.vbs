Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, cn 
Dim itemtype

p1 = "HKLM\SOFTWARE\Microsoft\DrWatson\"

n = ws.RegRead(p1 & "WaveFile")
t = "Change the Application Error Wave Sound"
cn = InputBox("Type New WaveFile (.wav) and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "WaveFile", cn
End If
