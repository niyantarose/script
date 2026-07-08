# ============================================================
# Project_19（発注EMSリスト）安全push
#   clasp push はローカルの中身でオンライン全体を上書きするため、
#   共同作業者がApps Scriptエディタで編集したファイルが毎回巻き戻っていた。
#   このスクリプトは push の前にオンライン編集を取り込んでから push する。
#
#   使い方: powershell -ExecutionPolicy Bypass -File tools\p19_safe_push.ps1
# ============================================================
$ErrorActionPreference = 'Stop'

# 共同作業者がオンラインで編集するファイル（ここに載っているものは自動取り込み）
$remoteOwned = @('インボイス.js', 'P-touch.gs.js', 'P-touch-CSV‗UTF16.js', '無題.js')

$p19 = (Resolve-Path (Join-Path $PSScriptRoot '..\Project_19')).Path

# 0) Project_19に未コミットの変更があれば中止（先にコミットしてから実行）
Push-Location $p19
$dirty = git status --porcelain -- .
if ($dirty) {
  Pop-Location
  Write-Host "Project_19に未コミットの変更があります。先にコミットしてください：`n$dirty"
  exit 1
}
Pop-Location

# 1) オンライン版を一時フォルダへpull
$tmp = Join-Path $env:TEMP ("p19_pull_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
Copy-Item (Join-Path $p19 '.clasp.json') $tmp
Push-Location $tmp
clasp pull | Out-Null
Pop-Location

# 2) 共同作業者ファイルの差分を取り込んでコミット
$absorbed = @()
$warn = @()
Get-ChildItem $tmp -File | Where-Object Name -ne '.clasp.json' | ForEach-Object {
  $local = Join-Path $p19 $_.Name
  $differs = (-not (Test-Path $local)) -or
    ((Get-FileHash $_.FullName).Hash -ne (Get-FileHash $local).Hash)
  if (-not $differs) { return }
  if ($remoteOwned -contains $_.Name) {
    Copy-Item $_.FullName $local -Force
    $absorbed += $_.Name
  } else {
    $warn += $_.Name
  }
}
Remove-Item -Recurse -Force $tmp

if ($absorbed) {
  Push-Location $p19
  git add -- $absorbed
  git commit -m "sync(p19): オンライン編集分を取り込み ($($absorbed -join ', '))"
  Pop-Location
  Write-Host "オンライン編集を取り込みました: $($absorbed -join ', ')"
}
if ($warn) {
  Write-Host "⚠ オンライン側にだけ変更があるファイル（自動取り込み対象外）: $($warn -join ', ')"
  Write-Host "  ローカルの変更をこれからpushで上書きします。取り込みが必要なら中止してください。"
  $ans = Read-Host "続行する？ (y/N)"
  if ($ans -ne 'y') { Write-Host '中止しました。'; exit 1 }
}

# 3) push
Push-Location $p19
clasp push -f
Pop-Location
Write-Host "✅ 安全push完了（取り込み: $($absorbed.Count)件）"
