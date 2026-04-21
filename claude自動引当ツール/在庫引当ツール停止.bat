@echo off
chcp 65001 > nul
echo 在庫引当ツール (port 5001) を停止します...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5001 " ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a > nul 2>&1
    echo PID %%a を停止しました
)
echo 完了
timeout /t 2 > nul
