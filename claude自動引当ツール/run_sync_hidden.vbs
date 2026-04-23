Set objShell = CreateObject("WScript.Shell")
strDir = "C:\Users\Owner\Desktop\script\claude自動引当ツール"
strCmd = "cmd /c cd /d """ & strDir & """ && python sync_agent.py >> sync_agent.log 2>&1"
' 0 = SW_HIDE: ウィンドウ完全非表示
objShell.Run strCmd, 0, False
