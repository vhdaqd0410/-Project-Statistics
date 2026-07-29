' AI后期剪辑提成表生成工具 - 无黑窗启动器
' 双击此文件直接启动GUI
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
guiPath = dir & "\ai_commission_gui.py"

If Not fso.FileExists(guiPath) Then
    MsgBox "找不到程序文件:" & vbCrLf & guiPath, 48, "错误"
    WScript.Quit 1
End If

On Error Resume Next
WshShell.Run "powershell -WindowStyle Hidden -Command " & Chr(34) & "& python '" & guiPath & "'" & Chr(34), 0, False
If Err.Number = 0 Then WScript.Quit 0
On Error GoTo 0

WshShell.Run "python """ & guiPath & """", 0, False
