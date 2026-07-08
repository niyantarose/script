# ============================================================
# Project_19（発注EMSリスト）安全push v3 — 全ファイル保護版
#   前回push時の状態(tools/p19_sync_state.json)を基準にした3方向同期:
#     ・オンラインだけ変わったファイル → リポジトリへ取り込み（syncコミット）
#     ・ローカルだけ変わったファイル → pushで反映
#     ・両方変わったファイル → どちらを残すか対話で確認
#   どのファイルが相手でも、共同作業者のオンライン編集は消えない。
#
#   使い方: powershell -ExecutionPolicy Bypass -File tools\p19_safe_push.ps1
# ============================================================
$ErrorActionPreference = 'Stop'

$p19 = (Resolve-Path (Join-Path $PSScriptRoot '..\Project_19')).Path
$statePath = Join-Path $PSScriptRoot 'p19_sync_state.json'

function Get-ContentHash([string]$path) { (Get-FileHash $path -Algorithm SHA256).Hash }

# 0) 未コミットの変更があれば中止（先にコミットしてから実行）
Push-Location $p19
$dirty = git status --porcelain -- .
Pop-Location
if ($dirty) {
  Write-Host "Project_19に未コミットの変更があります。先にコミットしてください：`n$dirty"
  exit 1
}

# 1) オンライン版を一時フォルダへpull
$tmp = Join-Path $env:TEMP ("p19_pull_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
Copy-Item (Join-Path $p19 '.clasp.json') $tmp
Push-Location $tmp
clasp pull | Out-Null
Pop-Location

# 2) 前回push時の状態（基準）を読む
$base = @{}
if (Test-Path $statePath) {
  (Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties |
    ForEach-Object { $base[$_.Name] = $_.Value }
} else {
  Write-Host "⚠ 同期状態ファイルが無い初回実行です。差分のあるファイルは対話で確認します。"
}

$absorbed = @()
$pushing = @()
$conflicts = @()

Get-ChildItem $tmp -File | Where-Object Name -ne '.clasp.json' | ForEach-Object {
  $name = $_.Name
  $localPath = Join-Path $p19 $name

  # オンラインにしかない新規ファイルは無条件で取り込む（pushで消してしまわないように）
  if (-not (Test-Path $localPath)) {
    Copy-Item $_.FullName $localPath -Force
    $absorbed += $name
    return
  }

  $onlineHash = Get-ContentHash $_.FullName
  $localHash = Get-ContentHash $localPath
  if ($onlineHash -eq $localHash) { return }

  $baseHash = $base[$name]
  if ($baseHash -and $onlineHash -eq $baseHash) {
    $pushing += $name          # ローカルだけ変更 → pushで反映（通常の開発フロー）
  } elseif ($baseHash -and $localHash -eq $baseHash) {
    Copy-Item $_.FullName $localPath -Force
    $absorbed += $name         # オンラインだけ変更 → 取り込み
  } else {
    $conflicts += [pscustomobject]@{ Name = $name; Online = $_.FullName; Local = $localPath }
  }
}

# ローカルにあってオンラインに無いファイルの案内
Get-ChildItem $p19 -File | Where-Object Name -ne '.clasp.json' | ForEach-Object {
  if (-not (Test-Path (Join-Path $tmp $_.Name))) {
    Write-Host "ℹ オンラインに無い「$($_.Name)」はpushで復活します（オンライン側で意図的に削除した場合はローカルからも削除してください）"
  }
}

# 3) 両方変わったファイルはどちらを残すか確認
foreach ($c in $conflicts) {
  Write-Host "⚠ 「$($c.Name)」はオンラインとローカルの両方が変更されています。"
  $ans = Read-Host "  どちらを残す？ [o]=オンライン版を取り込む / [l]=ローカル版でpush / それ以外=中止"
  if ($ans -eq 'o') { Copy-Item $c.Online $c.Local -Force; $absorbed += $c.Name }
  elseif ($ans -eq 'l') { $pushing += $c.Name }
  else { Write-Host '中止しました。'; Remove-Item -Recurse -Force $tmp; exit 1 }
}

# 4) 取り込み分をコミット
if ($absorbed) {
  Push-Location $p19
  git add -- $absorbed
  git commit -m "sync(p19): オンライン編集分を取り込み ($($absorbed -join ', '))"
  Pop-Location
  Write-Host "オンライン編集を取り込みました: $($absorbed -join ', ')"
}
if ($pushing) {
  Write-Host "ローカルの変更をpushします: $($pushing -join ', ')"
}

# 5) push
Push-Location $p19
clasp push -f
Pop-Location

# 6) push後の状態を記録して次回の基準にする
$state = [ordered]@{}
Get-ChildItem $p19 -File | Where-Object Name -ne '.clasp.json' | Sort-Object Name | ForEach-Object {
  $state[$_.Name] = Get-ContentHash $_.FullName
}
$enc = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($statePath, ($state | ConvertTo-Json), $enc)
Push-Location $p19
git add -- $statePath
git commit -m "chore(p19): 同期状態を更新（安全pushの基準）" | Out-Null
Pop-Location

Remove-Item -Recurse -Force $tmp
Write-Host "✅ 安全push完了（オンライン取り込み $($absorbed.Count)件 / ローカル反映 $($pushing.Count)件）"
