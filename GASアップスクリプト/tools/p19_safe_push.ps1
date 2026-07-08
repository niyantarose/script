# ============================================================
# Project_19（発注EMSリスト）安全push
#   実体は全プロジェクト共通版 gas_safe_push.ps1（3方向同期）。
#   使い方: powershell -ExecutionPolicy Bypass -File tools\p19_safe_push.ps1
# ============================================================
& (Join-Path $PSScriptRoot 'gas_safe_push.ps1') -Project 'Project_19'
exit $LASTEXITCODE
