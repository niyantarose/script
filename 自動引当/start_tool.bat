@echo off
cd /d %~dp0
pwsh -ExecutionPolicy Bypass -File "%~dp0start_tool.ps1"
