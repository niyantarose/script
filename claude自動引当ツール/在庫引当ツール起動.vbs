Dim shell, fso, dir, logFile
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

dir = fso.GetParentFolderName(WScript.ScriptFullName)
logFile = dir & "\server.log"

' すでに起動中か確認
Dim checkCmd
checkCmd = "netstat -ano | findstr :5001"
Dim result
result = shell.Run("cmd /c " & checkCmd & " > """ & dir & "\__check.tmp"" 2>&1", 0, True)

Dim tmpFile, alreadyRunning
alreadyRunning = False
If fso.FileExists(dir & "\__check.tmp") Then
    Set tmpFile = fso.OpenTextFile(dir & "\__check.tmp", 1)
    Dim content
    If Not tmpFile.AtEndOfStream Then
        content = tmpFile.ReadAll
    Else
        content = ""
    End If
    tmpFile.Close
    fso.DeleteFile dir & "\__check.tmp"
    If InStr(content, "5001") > 0 Then
        alreadyRunning = True
    End If
End If

If alreadyRunning Then
    shell.Run "cmd /c start http://127.0.0.1:5001", 0, False
    WScript.Quit
End If

' Pythonを探す
Dim pythonPath
pythonPath = ""
Dim candidates(3)
candidates(0) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\python.exe"
candidates(1) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\python.exe"
candidates(2) = "python"

Dim i
For i = 0 To 1
    If fso.FileExists(candidates(i)) Then
        pythonPath = candidates(i)
        Exit For
    End If
Next
If pythonPath = "" Then pythonPath = "python"

' バックグラウンドでFlask起動（ウィンドウなし）
Dim cmd
cmd = """" & pythonPath & """ """ & dir & "\start_server.py"" >> """ & logFile & """ 2>&1"
shell.Run "cmd /c " & cmd, 0, False

' 同期エージェントも起動（port 5050でHTTPトリガーサーバー）
Dim syncLog, syncCmd
syncLog = dir & "\sync_agent.log"
syncCmd = """" & pythonPath & """ """ & dir & "\sync_agent.py"" --server >> """ & syncLog & """ 2>&1"
shell.Run "cmd /c " & syncCmd, 0, False

' 少し待ってブラウザを開く
WScript.Sleep 2000
shell.Run "cmd /c start http://127.0.0.1:5001", 0, False
