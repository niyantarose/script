<?php
/**
 * yahoo_execute_api.php
 *
 * 役割:
 *  - preprocess が保存した /var/www/tmp/yahoo_images/<code>/images.zip を使って
 *    Yahooへ「全削除→ZIPアップロード→submit」を実行する
 *  - Authorization: Bearer は GAS 側で管理（VPSは受け取って使うだけ）
 */

header('Content-Type: application/json; charset=utf-8');

$sellerId = 'niyantarose';
$baseRoot = '/var/www/tmp/yahoo_images';

// --------------------------------------------------
// Bearer check
// --------------------------------------------------
$auth = $_SERVER['HTTP_AUTHORIZATION'] ?? '';
if (!preg_match('/^Bearer\s+(.+)$/', $auth, $m)) {
  http_response_code(401);
  echo json_encode(['ok'=>false,'error'=>'invalid bearer'], JSON_UNESCAPED_UNICODE);
  exit;
}
$accessToken = trim($m[1]);
if (strlen($accessToken) < 50) {
  http_response_code(401);
  echo json_encode(['ok'=>false,'error'=>'invalid bearer'], JSON_UNESCAPED_UNICODE);
  exit;
}

// --------------------------------------------------
// input
// --------------------------------------------------
$raw = file_get_contents('php://input');
$payload = json_decode($raw, true);
$itemCode = $payload['item_code'] ?? '';

if (!$itemCode || !preg_match('/^[A-Za-z0-9._-]+$/', $itemCode)) {
  http_response_code(400);
  echo json_encode(['ok'=>false,'error'=>'invalid item_code','raw'=>$payload], JSON_UNESCAPED_UNICODE);
  exit;
}

$zipPath = "{$baseRoot}/{$itemCode}/images.zip";
if (!file_exists($zipPath) || filesize($zipPath) <= 0) {
  http_response_code(404);
  echo json_encode(['ok'=>false,'error'=>'images.zip not found','zip'=>$zipPath], JSON_UNESCAPED_UNICODE);
  exit;
}

// --------------------------------------------------
// lock (同じitem_codeの同時実行防止)
// --------------------------------------------------
$lockFile = "/tmp/yahoo_exec_lock_{$itemCode}.lock";
$lockFp = fopen($lockFile, 'c');
if (!$lockFp || !flock($lockFp, LOCK_EX | LOCK_NB)) {
  http_response_code(429);
  echo json_encode(['ok'=>false,'error'=>'locked (already running?)','item_code'=>$itemCode], JSON_UNESCAPED_UNICODE);
  exit;
}

$runId = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
$workDir = "/tmp/yahoo_exec_runs/{$runId}/{$itemCode}";
@mkdir($workDir, 0775, true);

try {
  // ここで unzip + preprocess を入れたくなったら workDir に展開して加工→再zip する
  // ただし v28 の設計では、GAS側でZIP生成してるので、execute側は基本そのまま送ればOK

  // --------------------------------------------------
  // helper: curl POST
  // --------------------------------------------------
  $curlPost = function($url, $postFields, $accessToken, $isMultipart = false) {
    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, $postFields);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_TIMEOUT, 180);

    $headers = ['Authorization: Bearer ' . $accessToken];
    if ($isMultipart) {
      // multipart は curl が境界を作るので Content-Type を固定指定しない
    } else {
      $headers[] = 'Content-Type: application/x-www-form-urlencoded';
    }
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);

    $res = curl_exec($ch);
    $err = curl_error($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    return [$code, $res, $err];
  };

  // --------------------------------------------------
  // 1) delete all images
  // --------------------------------------------------
  [$dCode, $dBody, $dErr] = $curlPost(
    'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage',
    http_build_query([
      'seller_id' => $sellerId,
      'item_code' => $itemCode, // ★全削除
    ]),
    $accessToken,
    false
  );

  // delete がコケたら “ゴミ残り” リスクが出るので止める（安全運用）
  if ($dCode < 200 || $dCode >= 300) {
    http_response_code(502);
    echo json_encode([
      'ok' => false,
      'phase' => 'delete',
      'http' => $dCode,
      'err' => $dErr,
      'body' => $dBody,
    ], JSON_UNESCAPED_UNICODE);
    exit;
  }

  // --------------------------------------------------
  // 2) upload zip
  // --------------------------------------------------
  $uploadUrl = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImagePack?seller_id={$sellerId}";
  $cfile = new CURLFile($zipPath, 'application/zip', 'images.zip');

  [$uCode, $uBody, $uErr] = $curlPost(
    $uploadUrl,
    ['file' => $cfile],
    $accessToken,
    true
  );

  if ($uCode < 200 || $uCode >= 300) {
    http_response_code(502);
    echo json_encode([
      'ok' => false,
      'phase' => 'upload',
      'http' => $uCode,
      'err' => $uErr,
      'body' => $uBody,
    ], JSON_UNESCAPED_UNICODE);
    exit;
  }

  // --------------------------------------------------
  // 3) submit
  // --------------------------------------------------
  [$sCode, $sBody, $sErr] = $curlPost(
    'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem',
    http_build_query([
      'seller_id' => $sellerId,
      'item_code' => $itemCode,
    ]),
    $accessToken,
    false
  );

  if ($sCode < 200 || $sCode >= 300) {
    http_response_code(502);
    echo json_encode([
      'ok' => false,
      'phase' => 'submit',
      'http' => $sCode,
      'err' => $sErr,
      'body' => $sBody,
    ], JSON_UNESCAPED_UNICODE);
    exit;
  }

  echo json_encode([
    'ok' => true,
    'item_code' => $itemCode,
    'run_id' => $runId,
    'zip' => $zipPath,
    'delete' => ['http'=>$dCode],
    'upload' => ['http'=>$uCode],
    'submit' => ['http'=>$sCode],
  ], JSON_UNESCAPED_UNICODE);

} finally {
  // lock release
  if ($lockFp) {
    flock($lockFp, LOCK_UN);
    fclose($lockFp);
  }
  // workdir は今は使ってないから掃除してOK（必要なら残してデバッグ）
  // @exec("rm -rf " . escapeshellarg(dirname($workDir)));
}
