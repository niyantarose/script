# ============================================================
# GAS安全push（全プロジェクト共通版）
#   前回push時の状態(tools/sync_state/<Project>.json)を基準にした3方向同期:
#     ・オンラインだけ変わったファイル → リポジトリへ取り込み（syncコミット）
#     ・ローカルだけ変わったファイル → pushで反映
#     ・両方変わったファイル → どちらを残すか対話で確認
#   オンラインの新規ファイルは自動取り込み。オンラインで削除されたファイルは案内のみ。
#   初回実行（基準ファイルなし）では、オンライン版の内容がgit履歴に存在すれば
#   「ローカルが新しい」と自動判定し、未知の内容なら対話で確認する。
#
#   push後、固定デプロイ(Webアプリ)が1つだけあるプロジェクトは本番デプロイまで自動実行する。
#   （/exec URL は「デプロイした版」を実行する仕様で、push だけでは反映されないため）
#
#   使い方: powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_19
#           powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_19 -SkipDeploy
# ============================================================
param(
  [Parameter(Mandatory = $true)][string]$Project,
  # Webアプリの本番デプロイまで自動で行う。反映したくないときだけ -SkipDeploy を付ける。
  [switch]$SkipDeploy
)
$ErrorActionPreference = 'Stop'

$dir = Join-Path (Split-Path $PSScriptRoot -Parent) $Project
if (-not (Test-Path (Join-Path $dir '.clasp.json'))) {
  Write-Host "$Project に .clasp.json が見つかりません。"
  exit 1
}
$stateDir = Join-Path $PSScriptRoot 'sync_state'
if (-not (Test-Path $stateDir)) { New-Item -ItemType Directory $stateDir | Out-Null }
$statePath = Join-Path $stateDir ($Project + '.json')

function Get-ContentHash([string]$path) { (Get-FileHash $path -Algorithm SHA256).Hash }

# 0) 未コミットの変更があれば中止（先にコミットしてから実行）
Push-Location $dir
$dirty = git status --porcelain -- .
Pop-Location
if ($dirty) {
  Write-Host "$Project に未コミットの変更があります。先にコミットしてください：`n$dirty"
  exit 1
}

# 1) オンライン版を一時フォルダへpull
$tmp = Join-Path $env:TEMP ('gas_pull_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
Copy-Item (Join-Path $dir '.clasp.json') $tmp
Push-Location $tmp
clasp pull | Out-Null
Pop-Location

# 2) 前回push時の状態（基準）を読む
$base = @{}
$firstRun = -not (Test-Path $statePath)
if (-not $firstRun) {
  (Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties |
    ForEach-Object { $base[$_.Name] = $_.Value }
} else {
  Write-Host "ℹ $Project は初回実行です（基準ファイルを作成します）。"
}

$absorbed = @()
$pushing = @()
$conflicts = @()

Get-ChildItem $tmp -File | Where-Object Name -ne '.clasp.json' | ForEach-Object {
  $name = $_.Name
  $localPath = Join-Path $dir $name

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
    return
  }
  if ($baseHash -and $localHash -eq $baseHash) {
    Copy-Item $_.FullName $localPath -Force
    $absorbed += $name         # オンラインだけ変更 → 取り込み
    return
  }

  # 基準なし（初回）や基準と両方違う場合:
  # オンライン版の内容がgit履歴にあれば「既知の旧版＝ローカルが新しい」と判定
  Push-Location $dir
  $blob = git hash-object $_.FullName
  git cat-file -e $blob 2>$null
  $known = ($LASTEXITCODE -eq 0)
  Pop-Location
  if ($known) {
    $pushing += $name
  } else {
    $conflicts += [pscustomobject]@{ Name = $name; Online = $_.FullName; Local = $localPath }
  }
}

# ローカルにあってオンラインに無いファイルの案内
Get-ChildItem $dir -File | Where-Object Name -ne '.clasp.json' | ForEach-Object {
  if (-not (Test-Path (Join-Path $tmp $_.Name))) {
    Write-Host "ℹ オンラインに無い「$($_.Name)」はpushで復活します（オンライン側で意図的に削除した場合はローカルからも削除してください）"
  }
}

# 3) 判定できないファイルはどちらを残すか確認
foreach ($c in $conflicts) {
  Write-Host "⚠ 「$($c.Name)」はオンライン側に未知の変更があります（ローカル側も基準と異なる可能性）。"
  $ans = Read-Host "  どちらを残す？ [o]=オンライン版を取り込む / [l]=ローカル版でpush / それ以外=中止"
  if ($ans -eq 'o') { Copy-Item $c.Online $c.Local -Force; $absorbed += $c.Name }
  elseif ($ans -eq 'l') { $pushing += $c.Name }
  else { Write-Host '中止しました。'; Remove-Item -Recurse -Force $tmp; exit 1 }
}

# 4) 取り込み分をコミット
if ($absorbed) {
  Push-Location $dir
  git add -- $absorbed
  git commit -m "sync($Project): オンライン編集分を取り込み ($($absorbed -join ', '))"
  Pop-Location
  Write-Host "オンライン編集を取り込みました: $($absorbed -join ', ')"
}
if ($pushing) {
  Write-Host "ローカルの変更をpushします: $($pushing -join ', ')"
}

# 5) push
Push-Location $dir
clasp push -f
Pop-Location

# 6) push後の状態を記録して次回の基準にする
$state = [ordered]@{}
Get-ChildItem $dir -File | Where-Object Name -ne '.clasp.json' | Sort-Object Name | ForEach-Object {
  $state[$_.Name] = Get-ContentHash $_.FullName
}
$enc = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($statePath, ($state | ConvertTo-Json), $enc)
Push-Location $dir
git add -- $statePath
git commit -m "chore($Project): 同期状態を更新（安全pushの基準）" | Out-Null
Pop-Location

Remove-Item -Recurse -Force $tmp
Write-Host "✅ 安全push完了 [$Project]（オンライン取り込み $($absorbed.Count)件 / ローカル反映 $($pushing.Count)件）"

# 7) Webアプリの本番デプロイ（push だけでは /exec に反映されないため）
#    /exec URL は「デプロイした版」を実行する仕様で、HEAD を自動追従しない。
#    HEAD 追従するのは /dev だが編集権限が要るので拡張機能からは使えない。
#    → 固定デプロイが1つだけあるプロジェクトは、ここで自動的に版を上げる。
if (-not $SkipDeploy) {
  Push-Location $dir
  try {
    $deployOutput = & npx clasp deployments 2>&1 | Out-String
    # 「- <deploymentId> @<version> - <説明>」の形式。@HEAD（テストデプロイ）は対象外。
    # @(...) で必ず配列にする。1件だと文字列になり $versioned[0] が先頭1文字になるため。
    $versioned = @(
      [regex]::Matches($deployOutput, '(?m)^-\s+(?<id>AK\S+)\s+@(?<ver>\d+)') |
        ForEach-Object { $_.Groups['id'].Value } | Select-Object -Unique
    )

    if ($versioned.Count -eq 0) {
      Write-Host "ℹ $Project に固定デプロイはありません（デプロイ不要）"
    }
    elseif ($versioned.Count -gt 1) {
      # どれを更新すべきか機械的に決められないので、勝手に本番を書き換えない
      Write-Host "⚠ $Project に固定デプロイが複数あります。自動デプロイをスキップしました:"
      $versioned | ForEach-Object { Write-Host "    $_" }
      Write-Host "  必要なら手動で: npx clasp deploy -i <デプロイID> -d `"説明`""
    }
    else {
      $deployId = $versioned[0]
      $desc = "auto: safe_push $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
      Write-Host "本番デプロイを実行します [$Project] $deployId"
      $result = & npx clasp deploy -i $deployId -d $desc 2>&1 | Out-String
      Write-Host $result.Trim()
      if ($result -notmatch 'Deployed') {
        Write-Host "⚠ デプロイ結果に 'Deployed' がありません。反映されたか確認してください。"
      }
    }
  } catch {
    Write-Host "⚠ 自動デプロイでエラー: $($_.Exception.Message)"
    Write-Host "  手動で Project フォルダから: npx clasp deploy -i <デプロイID> -d `"説明`""
  } finally {
    Pop-Location
  }
}
