Option Explicit

' =================================================================================================
' ToggleNetConnection.vbs - Bill James - 13 Dec 2004
' Utility to Enable or Disable a network connection.
' Settings to edit below:
'   sConnName: The name of the connection as viewed in the Windows Network Connections folder
'   sNetFldr:  The name of the Windows Network Connections folder.  In Windows 2000 this is called
'              "Network and Dial-up Connections", in Windows XP it is called "Network Connections".
' =================================================================================================

' ========================= SETTINGS =========================
sConnName = "Wireless Network Connection"
sNetFldr = "Network Connections"
' ============================================================



Dim sConnName, sNetFldr
Dim sConnect, sDisconnect
' Verbs - The right click context menu choices (& indicates the shortcut key)
'         C&onnect/Disc&onnect are used for dial-up and VPN connections, 
'         En&able/Disa&ble are for NIC connections.
'sConnect = "C&onnect"              'Dialup or VPN Connect
'sDisconnect = "Disc&onnect"        'Dialup or VPN Disconnect
sConnect = "En&able"               'NIC or Wireless Enable
sDisconnect = "Disa&ble"           'NIC or Wireless Disable

Dim FldrItem, oNetConn
For Each FldrItem in CreateObject("Shell.Application").Namespace(3).Items
  If FldrItem.Name = sNetFldr Then
    Set oNetConn = FldrItem.GetFolder
    Exit For
  End If
next

If Not IsObject(oNetConn) Then
  MsgBox "Couldn't find '" & sNetFldr & "' folder"
  WScript.Quit
End If

Dim oLanConnection
For Each FldrItem in oNetConn.items
  If LCase(FldrItem.Name) = LCase(sConnName) Then
    Set oLanConnection = FldrItem
    Exit For
  End If
next

If Not IsObject(oLanConnection) Then
  MsgBox "Couldn't find '" & sConnName & "' item"
  WScript.Quit
End If

Dim bEnabled
bEnabled = True
Dim oVerb, oConnectVerb, oDisconnectVerb
For Each oVerb in oLanConnection.verbs
  If oVerb.Name = sConnect Then 
    Set oConnectVerb = oVerb 
    bEnabled = False
  End If
  If oVerb.Name = sDisconnect Then 
    Set oDisconnectVerb = oVerb 
  End If
Next

Dim sMsg
If bEnabled Then
  oDisconnectVerb.DoIt
  sMsg = sConnName & " has been disabled."
Else
  oConnectVerb.DoIt
  sMsg = sConnName & " has been enabled."
End If

' A sleep is needed, might be able to use less
WScript.Sleep 2000 

CreateObject("WScript.Shell").Popup sMsg, 3, "Toggle Network Connection"
