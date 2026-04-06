<?php
header('Content-Type: application/json');

// ==================================================
// 受信
// ==================================================
$raw  = file_get_contents('php://input');
$data = json_decode($raw, true);

if (!$data) {
  http_response_code(400);
  echo json_encode(['ok'=>false,'error'=>'invalid json','raw'=>$raw]);
  exit;
}

$itemCode = $data['item_code'] ?? null;
$images   = $data['images'] ?? null;

if (!$itemCode || !is_array($images)) {
  http_response_code(400);
  echo json_encode(['ok'=>false,'error'=>'missing item_code or images','parsed'=>$data]);
  exit;
}

// ==================================================
// 作業ディレクトリ作成
// ==================================================
$runId   = date('Ymd_His') . '_' . substr(bin2hex(random_bytes(4)), 0, 8);
$baseDir = "/tmp/yahoo_zip_runs/$runId";
$workDir = "$baseDir/work";

if (!mkdir($workDir, 0775, true)) {
  http_response_code(500);
  echo json_encode(['ok'=>false,'error'=>'failed to create workdir']);
  exit;
}

// ==================================================
// DL
// ==================================================
$downloaded = [];
$errors     = [];

$idx = 0;
foreach ($images as $url) {
  $idx++;

  // 仮ファイル名（後で正式リネーム）
  $dst = sprintf('%s/%s_%02d.jpg', $workDir, $itemCode, $idx);

  // URLじゃない場合はスキップ（今はA/B/C対策）
  if (!preg_match('#^https?://#i', $url)) {
    $errors[] = "skip(non-url): $url";
    continue;
  }

  $ch = curl_init($url);
  curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_TIMEOUT        => 30,
    CURLOPT_USERAGENT      => 'NiYANTA-ROSE Yahoo Image Bot',
  ]);

  $bin  = curl_exec($ch);
  $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
  $err  = curl_error($ch);
  curl_close($ch);

  if ($bin === false || $code !== 200) {
    $errors[] = "dl failed ($code): $url $err";
    continue;
  }

  file_put_contents($dst, $bin);
  $downloaded[] = basename($dst);
}

// ==================================================
// デバッグ保存
// ==================================================
file_put_contents(
  "$baseDir/debug.json",
  json_encode([
    'item_code'   => $itemCode,
    'images'      => $images,
    'downloaded'  => $downloaded,
    'errors'      => $errors,
    'workdir'     => $workDir,
  ], JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE)
);

// ==================================================
// 応答
// ==================================================
echo json_encode([
  'ok'         => true,
  'run_id'     => $runId,
  'workdir'    => $workDir,
  'downloaded' => $downloaded,
  'errors'     => $errors,
]);
exit;


