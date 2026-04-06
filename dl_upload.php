<?php
require_once __DIR__ . '/_vps_secret.php';

// ★一時デバッグ（確認後すぐ消す）
file_put_contents('/tmp/dl_upload_debug.log', json_encode([
  'ts' => date('Y-m-d H:i:s'),
  'header' => _vps_get_header('X-VPS-Secret'),
  'post_secret' => $_POST['secret'] ?? '(none)',
  'expected_len' => strlen(_vps_read_expected_secret()),
  'got_header_len' => strlen(_vps_get_header('X-VPS-Secret')),
  'got_post_len' => strlen($_POST['secret'] ?? ''),
  'match_header' => hash_equals(_vps_read_expected_secret(), _vps_get_header('X-VPS-Secret')),
  'match_post' => isset($_POST['secret']) && hash_equals(_vps_read_expected_secret(), $_POST['secret']),
], JSON_PRETTY_PRINT) . "\n", FILE_APPEND);

require_secret_or_403();



require_once __DIR__ . '/_vps_secret.php';
require_secret_or_403();

header('Content-Type: application/json; charset=utf-8');

function respond($arr, $code = 200) {
  http_response_code($code);
  echo json_encode($arr, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
  exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
  respond(['ok'=>false,'error'=>'METHOD_NOT_ALLOWED'], 405);
}

if (!isset($_FILES['zip'])) {
  respond(['ok'=>false,'error'=>'zip file missing','files'=>array_keys($_FILES ?? [])], 400);
}

$f = $_FILES['zip'];
if (($f['error'] ?? UPLOAD_ERR_OK) !== UPLOAD_ERR_OK) {
  respond(['ok'=>false,'error'=>'UPLOAD_ERR','code'=>$f['error']], 400);
}

$name = $_POST['name'] ?? '';
$name = is_string($name) ? trim($name) : '';
if ($name === '') {
  // name未指定なら元ファイル名
  $name = basename($f['name'] ?? 'upload.zip');
}
$name = basename($name);

// 拡張子チェック（安全側）
if (!preg_match('/\.zip$/i', $name)) {
  respond(['ok'=>false,'error'=>'name must end with .zip','name'=>$name], 400);
}

// 置き場
$dlDir = '/var/www/html/dl';
if (!is_dir($dlDir)) {
  if (!@mkdir($dlDir, 0775, true)) respond(['ok'=>false,'error'=>'mkdir dl failed'], 500);
}
$dst = $dlDir . '/' . $name;

// 上書きOK
if (!move_uploaded_file($f['tmp_name'], $dst)) {
  respond(['ok'=>false,'error'=>'move_uploaded_file failed'], 500);
}

// パーミッション整える（www-dataで読めればOK）
@chmod($dst, 0664);

respond([
  'ok' => true,
  'saved' => $dst,
  'url' => 'https://img.niyantarose.com/dl/' . rawurlencode($name),
  'size' => filesize($dst),
], 200);
