'Script to modify the MRU list for Start>Run,  Steve Yandl, Sept 11, 2000
Dim shell, dict, path, Ordering, Counter, TLimit, Letter, RunLine, k, i, s, DumpIt, Purge, Zap, Alphabet, Keepers, l, Newref
Dim TheLine
Set shell=CreateObject("Wscript.Shell")
Set dict = CreateObject("Scripting.Dictionary")
path="HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
On Error Resume Next
'Find out what is presently there and populate the dictionary object with the values
Ordering=shell.RegRead(path&"\MRUList")
If Err.Number<>0 Then MsgBox "There are no entries to clear":Set shell=Nothing:Set dict=Nothing:WScript.Quit
If Ordering="" Then MsgBox "There are no entries to clear":Set shell=Nothing:Set dict=Nothing:WScript.Quit
Counter=0
TLimit=Len(Ordering)
Do While Counter<TLimit
Counter=Counter+1
Letter=Mid(Ordering,Counter,1)
RunLine=shell.RegRead(path&"\"&Letter)
dict.add Letter, RunLine
Loop
    k=dict.keys
    For i = 0 To dict.Count -1 ' Iterate the array.
    TheLine=Left(dict.Item(k(i)),Len(dict.Item(k(i)))-2)
    s = s &k(i)&" >>   "& TheLine &vbCRLF ' Create return string.    
    Next
'Let the user pick items to dump
DumpIt=InputBox("ENTER INDEX LETTERS FOR ANY RUN LINES TO DELETE (Reboot after OK)"&vbCRLF&_
vbCRLF&"Index >> Run Line"&vbCRLF&s,"Kill These")
Purge=Len(DumpIt)
'Delete bum keys or escape program
    If Purge>0 Then
    For i = 0 To dict.Count -1 ' Iterate the array.
    shell.RegDelete path&"\"&k(i)        
    Next
    Else
    Set dict=Nothing
    Set shell=Nothing
    WScript.Quit
    End If
'Dump entries in the record we don't want back in the registry
    For j=1 To Purge
    Zap=Mid(DumpIt,j,1)
    dict.Remove Zap
    Next
'Reset the MRU list for remaining values and write them back to the registry
Alphabet="abcdefghijklmnopqurstvwxyz"
Keepers=Left(Alphabet,(TLimit-Purge))
shell.RegWrite path&"\MRUList",Keepers
'
   l=dict.keys
   For i = 0 To dict.Count-1
   Newref=Mid(Keepers,i+1,1)
   shell.RegWrite path&"\"&Newref,dict.Item(l(i))
   Next
Set dict=Nothing
Set shell=Nothing