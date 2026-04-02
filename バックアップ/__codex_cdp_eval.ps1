param(
  [Parameter(Mandatory = $true)][string]$TargetUrl,
  [Parameter(Mandatory = $true)][string]$Expression,
  [switch]$NewTarget
)

$ErrorActionPreference = 'Stop'
$baseUrl = 'http://127.0.0.1:9222'

if ($NewTarget) {
  $openUrl = "$baseUrl/json/new?" + [Uri]::EscapeDataString($TargetUrl)
  try {
    Invoke-WebRequest -UseBasicParsing -Method Put -Uri $openUrl | Out-Null
    Start-Sleep -Milliseconds 300
  } catch {
    # ignore if target already exists or new target failed transiently
  }
}

$targets = ((Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/json/list").Content | ConvertFrom-Json)
$target = $targets | Where-Object { $_.url -eq $TargetUrl } | Select-Object -First 1
if (-not $target) {
  throw "Target not found: $TargetUrl"
}

$ws = [System.Net.WebSockets.ClientWebSocket]::new()
$cts = [System.Threading.CancellationTokenSource]::new()
$ws.ConnectAsync([Uri]$target.webSocketDebuggerUrl, $cts.Token).GetAwaiter().GetResult()

function Send-CdpMessage {
  param([System.Net.WebSockets.ClientWebSocket]$Socket, [int]$Id, [string]$Method, [hashtable]$Params)
  $payload = @{ id = $Id; method = $Method; params = $Params } | ConvertTo-Json -Compress -Depth 20
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
  $segment = [System.ArraySegment[byte]]::new($bytes)
  $Socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).GetAwaiter().GetResult()
}

function Read-CdpMessage {
  param([System.Net.WebSockets.ClientWebSocket]$Socket)
  $buffer = New-Object byte[] 65536
  $builder = [System.Text.StringBuilder]::new()
  do {
    $segment = [System.ArraySegment[byte]]::new($buffer)
    $result = $Socket.ReceiveAsync($segment, $cts.Token).GetAwaiter().GetResult()
    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
      throw 'WebSocket closed'
    }
    $builder.Append([System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)) | Out-Null
  } while (-not $result.EndOfMessage)
  return $builder.ToString() | ConvertFrom-Json
}

try {
  Send-CdpMessage -Socket $ws -Id 1 -Method 'Runtime.enable' -Params @{}
  do {
    $msg = Read-CdpMessage -Socket $ws
  } until ($msg.id -eq 1)

  Send-CdpMessage -Socket $ws -Id 2 -Method 'Runtime.evaluate' -Params @{
    expression = $Expression
    awaitPromise = $true
    returnByValue = $true
    userGesture = $true
  }

  while ($true) {
    $msg = Read-CdpMessage -Socket $ws
    if ($msg.id -ne 2) { continue }
    if ($msg.error) {
      throw ($msg.error | ConvertTo-Json -Compress)
    }
    if ($msg.result.exceptionDetails) {
      throw ($msg.result.exceptionDetails | ConvertTo-Json -Depth 20 -Compress)
    }
    $msg.result.result | ConvertTo-Json -Depth 20 -Compress
    break
  }
} finally {
  if ($ws.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
    $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, 'done', $cts.Token).GetAwaiter().GetResult()
  }
  $ws.Dispose()
  $cts.Dispose()
}

