Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
guiPath = dir & "\ai_commission_gui.py"

If Not fso.FileExists(guiPath) Then
    MsgBox "Cannot find: " & guiPath, 48, "Error"
    WScript.Quit 1
End If

' 自动检测：优先 pythonw（无控制台），其次 python
cmd = "pythonw.exe """ & guiPath & """"
WshShell.Run cmd, 0, False
