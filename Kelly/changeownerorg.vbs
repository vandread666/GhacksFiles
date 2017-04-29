'ChgOwnOrg.vbs - Change Win9x Registered Owner/Organization.
'© Bill James - billjames.geo@yahoo.com - rev 29 Oct 1999.
'Modified for Windows XP and Vista

Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, g, cn, cg
Dim itemtype

p1 = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\"

n = ws.RegRead(p1 & "RegisteredOwner")
g = ws.RegRead(p1 & "RegisteredOrganization")
t = "Change Owner and Organization Utility"
cn = InputBox("Type new Owner and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "RegisteredOwner", cn
End If

cg = InputBox("Type new Organization and click OK.", t, g)
If cg <> "" Then
  ws.RegWrite p1 & "RegisteredOrganization", cg
End If