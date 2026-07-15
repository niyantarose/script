' 台湾CNシート×Yahoo実店舗の商品コード照合を非表示で実行する（タスクスケジューラ用）
' 結果: reports\recon_*.csv / 新規検出があれば reports\recon_NEW_*.csv / 推移は reports\recon_history.csv
Set objShell = CreateObject("WScript.Shell")
strDir = "C:\Users\Owner\Desktop\script\claude自動引当ツール"
strCmd = "cmd /c cd /d """ & strDir & """ && python scripts\yahoo_sheet_recon.py --max-wait 600 >> reports\recon_task.log 2>&1"
' 0 = SW_HIDE: ウィンドウ完全非表示
objShell.Run strCmd, 0, False
