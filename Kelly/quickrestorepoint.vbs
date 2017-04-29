Set IRP = getobject("winmgmts:\\.\root\default:Systemrestore")
MYRP = IRP.createrestorepoint ("Quick Restore Point", 0, 100)