# ============================================================
# 編集前の同期: オンライン(Apps Script)の編集をローカルへ取り込んでコミットする
#   スクリプトは共同作業者の貼り付けやオンライン編集でいつでも変わっている前提。
#   どのプロジェクトでも「編集を始める前」にこれを実行してから修正すること。
#
#   使い方: powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_19
# ============================================================
param([Parameter(Mandatory = $true)][string]$Project)
$ErrorActionPreference = 'Stop'

$dir = Join-Path (Split-Path $PSScriptRoot -Parent) $Project
if (-not (Test-Path (Join-Path $dir '.clasp.json'))) {
  Write-Host "$Project に .clasp.json が見つかりません。"
  exit 1
}

Push-Location $dir
try {
  # 未コミットの変更があると、pullで上書きして区別が付かなくなるので中止
  $dirty = git status --porcelain -- .
  if ($dirty) {
    Write-Host "$Project に未コミットの変更があります。先にコミットしてから実行してください："
    Write-Host $dirty
    exit 1
  }

  clasp pull | Out-Null
  $pulled = git status --porcelain -- .
  if ($pulled) {
    git add -A .
    git commit -m "sync($Project): オンライン編集分を取り込み"
    Write-Host "オンライン編集を取り込みました："
    Write-Host $pulled
  } else {
    Write-Host "$Project はオンラインとローカルが一致しています。"
  }
} finally {
  Pop-Location
}
