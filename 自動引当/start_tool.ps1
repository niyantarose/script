param(
    [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root ".venv\Scripts\python.exe"
$url = "http://127.0.0.1:5000"

if (-not (Test-Path $python)) {
    Write-Host "仮想環境が見つかりません: $python" -ForegroundColor Red
    Write-Host "先に README のセットアップ手順を実行してください。"
    exit 1
}

if (-not (Test-Path ".env") -and (Test-Path ".env.example")) {
    Copy-Item ".env.example" ".env"
}

function Test-AppReady {
    param([string]$Uri)

    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -TimeoutSec 1
        return $response.StatusCode -ge 200
    } catch {
        return $false
    }
}

$isRunning = Test-AppReady -Uri $url

if (-not $isRunning) {
    $escapedRoot = $root.Replace("'", "''")
    $serverCommand = "Set-Location '$escapedRoot'; & '$python' 'app.py'"
    Start-Process pwsh -ArgumentList "-NoExit", "-Command", $serverCommand | Out-Null

    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 500
        if (Test-AppReady -Uri $url) {
            $isRunning = $true
            break
        }
    }
}

if (-not $NoBrowser) {
    Start-Process $url
}

if ($isRunning) {
    Write-Host "自動引当ツールを開きました: $url" -ForegroundColor Green
} else {
    Write-Host "サーバーは起動処理を開始しました。数秒後にブラウザを確認してください: $url" -ForegroundColor Yellow
}
