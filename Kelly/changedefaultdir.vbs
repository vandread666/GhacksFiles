Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, cn
Dim itemtype

p1 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\"

n = ws.RegRead(p1 & "ProgramFilesDir")
t = "Change Default Installation Path Utility"
cn = InputBox("Type New Program Files Directory Path and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "ProgramFilesDir", cn
End If

