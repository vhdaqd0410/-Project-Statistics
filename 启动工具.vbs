Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
guiPath = dir & "\ai_commission_gui.py"
pyPath = "C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe"

If Not fso.FileExists(guiPath) Then
    MsgBox "Cannot find: " & guiPath, 48, "Error"
    WScript.Quit 1
End If

If Not fso.FileExists(pyPath) Then
    MsgBox "Cannot find Python: " & pyPath, 48, "Error"
    WScript.Quit 1
End If

cmd = """" & pyPath & """ """ & guiPath & """"
WshShell.Run cmd, 0, False
