@echo off
setlocal
cd /d %~dp0

echo =========================================
echo 画像分割カットツール かんたん開始
echo =========================================

after_python:
where py >nul 2>nul
if %errorlevel%==0 goto launch
where python >nul 2>nul
if %errorlevel%==0 goto launch

echo.
echo [ERROR] Python が見つかりません。
echo 先に Python 3.10+ をインストールしてください。
echo https://www.python.org/downloads/windows/
echo.
pause
exit /b 1

:launch
if not exist desktop_tool (
  echo [ERROR] desktop_tool フォルダが見つかりません。
  pause
  exit /b 1
)

call desktop_tool\run.bat
endlocal
