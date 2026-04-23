@echo off
cd /d %~dp0
if not exist desktop_tool (
  echo [ERROR] desktop_toolフォルダが見つかりません。
  pause
  exit /b 1
)
call desktop_tool\run.bat
