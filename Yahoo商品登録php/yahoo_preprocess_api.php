<?php
require_once __DIR__ . '/_vps_secret.php';
require_secret_or_403();


/**
 * yahoo_preprocess_api.php (FINAL)
 *
 * 役割:
 *  - GASから送られたZIPを item_code ごとの保存場所へコピーして置くだけ
 *  - Yahoo API は一切触らない
 *
 * 受信:
 *  - multipart/form-data
 *    - zip: file
 *    - item_codes: "AAA,BBB,CCC"
 *
 * 保存:
 *  /var/www/tmp/yahoo_images/<ITEM_CODE>/images.zip
 */

header('Content-Type: application/json; charset=utf-8');
umask(0002); // 664/775 を作りやすく

function respond($arr, $status = 200) {
  http_response_code($status);
  echo json_encode($arr, JSON_UNESCAPED_UNICODE);
  exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
  respond(['ok'=>false,'error'=>'method not allowed'], 405);
}

// ---- input check: zip ----
if (!isset($_FILES['zip']) || !is_array($_FILES['zip'])) {
  respond(['ok'=>false,'error'=>'zip file missing'], 400);
}
if (($_FILES['zip']['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
  respond(['ok'=>false,'error'=>'zip upload failed','file'=>$_FILES['zip']], 400);
}

// ---- item_codes parse (once) ----
$itemCodesRaw = $_POST['item_codes'] ?? '';
$itemCodesRaw = str_replace(['，','、',';','|'], ',', $itemCodesRaw); // 表記ゆれ吸収
$itemCodesRaw = preg_replace('/\s+/', '', $itemCodesRaw);           // 空白/改行除去

$itemCodes = array_values(array_filter(array_map('trim', explode(',', $itemCodesRaw))));
$itemCodes = array_values(array_unique($itemCodes));

if (empty($itemCodes)) {
  respond(['ok'=>false,'error'=>'item_codes missing','raw'=>$itemCodesRaw], 400);
}

// ---- dirs ----
$baseRoot   = '/var/www/tmp/yahoo_images';
$stagingDir = '/var/www/tmp/yahoo_preprocess_staging';

if (!is_dir($baseRoot) && !mkdir($baseRoot, 0775, true)) {
  respond(['ok'=>false,'error'=>'failed to create base dir','path'=>$baseRoot], 500);
}
if (!is_dir($stagingDir) && !mkdir($stagingDir, 0775, true)) {
  respond(['ok'=>false,'error'=>'failed to create staging dir','path'=>$stagingDir], 500);
}

// ---- stage zip once (IMPORTANT) ----
$runId = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
$stagingZip = $stagingDir . "/upload_{$runId}.zip";

// move_uploaded_file は “元の tmp” にしか効かない＆1回勝負なので、ここで確定させる
$tmp = $_FILES['zip']['tmp_name'];

$ok = @move_uploaded_file($tmp, $stagingZip);
if (!$ok) {
  // 環境によって move が失敗するケースがあるので copy フォールバック
  if (!@copy($tmp, $stagingZip)) {
    respond(['ok'=>false,'error'=>'failed to stage zip'], 500);
  }
}

$stagedSize = @filesize($stagingZip);
if ($stagedSize === false) $stagedSize = null;

// ---- save zip for each item_code ----
$saved  = [];
$errors = [];

foreach ($itemCodes as $code) {
  $code = trim($code);
  if ($code === '') continue;

  // パストラバーサル対策
  if (!preg_match('/^[A-Za-z0-9._-]+$/', $code)) {
    $errors[] = "invalid item_code: {$code}";
    continue;
  }

  $itemDir = "{$baseRoot}/{$code}";
  if (!is_dir($itemDir) && !mkdir($itemDir, 0775, true)) {
    $errors[] = "mkdir failed: {$code}";
    continue;
  }

  $dst = "{$itemDir}/images.zip";

  // ★ここが肝：stagingZip から copy（move_uploaded_file は絶対使わない）
  if (!@copy($stagingZip, $dst)) {
    $errors[] = "zip save failed: {$code}";
    continue;
  }

  @chmod($dst, 0664);
  $saved[] = $code;
}

// 後片付け（推奨）
@unlink($stagingZip);

respond([
  'ok'         => true,
  'run_id'     => $runId,
  'saved'      => $saved,
  'errors'     => $errors,
  'staged_size'=> $stagedSize,
]);
