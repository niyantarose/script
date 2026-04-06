<?php
header('Content-Type: application/json; charset=utf-8');

function jexit($arr, $code=200){
  http_response_code($code);
  echo json_encode($arr, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
  exit;
}

function read_request_json(){
  $raw = file_get_contents('php://input');
  if ($raw !== false && strlen(trim($raw)) > 0) {
    $j = json_decode($raw, true);
    if (is_array($j)) return $j;
  }

  // form fallback（GAS互換）
  $keys = ['json','payload','payload_json','request_json','data','body','items_json','files_json'];
  foreach($keys as $k){
    if (isset($_POST[$k]) && $_POST[$k] !== '') {
      $v = $_POST[$k];
      $j = json_decode($v, true);
      if (is_array($j)) return $j;
    }
  }

  if (!empty($_POST)) return $_POST;
  return null;
}

function safe_public_zip_name($s){
  $s = trim((string)$s);
  if ($s === '') return null;
  // 英数 _ - のみ
  if (!preg_match('/^[0-9A-Za-z_\-]+$/', $s)) return null;
  return $s;
}

$req = read_request_json();
if (!$req) jexit(['ok'=>false,'error'=>'Invalid JSON'], 400);

$publicZipName = safe_public_zip_name($req['publicZipName'] ?? '');
if (!$publicZipName) jexit(['ok'=>false,'error'=>'publicZipName invalid'], 400);

$files = $req['files'] ?? null;
if (!is_array($files) || count($files) === 0) jexit(['ok'=>false,'error'=>'files empty'], 400);

// ★保存先を dl に統一
$ZIP_DIR = '/var/www/html/dl';
$PUBLIC_BASE = 'https://img.niyantarose.com/dl';

if (!is_dir($ZIP_DIR)) {
  if (!@mkdir($ZIP_DIR, 0775, true)) jexit(['ok'=>false,'error'=>'zip dir create failed'], 500);
}

// ローカルパス許可（安全のため）
$ALLOWED_PREFIXES = [
  '/data/yahoo_images/',
  '/var/www/html/' // 念のため
];

$zipFile = $publicZipName . '.zip';
$zipPath = rtrim($ZIP_DIR, '/') . '/' . $zipFile;

// 既存があれば上書き
if (file_exists($zipPath)) @unlink($zipPath);

$zip = new ZipArchive();
if ($zip->open($zipPath, ZipArchive::CREATE) !== true) {
  jexit(['ok'=>false,'error'=>'zip open failed'], 500);
}

$missing = [];
$added = 0;

foreach($files as $f){
  if (!is_array($f)) continue;

  $localPath = (string)($f['localPath'] ?? '');
  $zipName   = (string)($f['zipName'] ?? '');

  if ($localPath === '' || $zipName === '') continue;

  // zipName の危険文字を排除（../ とか）
  if (strpos($zipName, '..') !== false) {
    $missing[] = ['zipName'=>$zipName,'localPath'=>$localPath,'reason'=>'zipName invalid'];
    continue;
  }

  // localPath は許可prefixのみ
  $okPrefix = false;
  foreach($ALLOWED_PREFIXES as $p){
    if (strncmp($localPath, $p, strlen($p)) === 0) { $okPrefix = true; break; }
  }
  if (!$okPrefix) {
    $missing[] = ['zipName'=>$zipName,'localPath'=>$localPath,'reason'=>'path not allowed'];
    continue;
  }

  if (!is_file($localPath)) {
    $missing[] = ['zipName'=>$zipName,'localPath'=>$localPath,'reason'=>'not found'];
    continue;
  }

  if (!$zip->addFile($localPath, $zipName)) {
    $missing[] = ['zipName'=>$zipName,'localPath'=>$localPath,'reason'=>'addFile failed'];
    continue;
  }

  $added++;
}

$zip->close();

if (!is_file($zipPath)) {
  jexit(['ok'=>false,'error'=>'zip not created'], 500);
}

jexit([
  'ok' => true,
  'zipUrl' => $PUBLIC_BASE . '/' . $zipFile,
  'zip' => $zipFile,
  'stats' => [
    'files_in_zip' => $added,
    'missing_count' => count($missing)
  ],
  'missing' => $missing
], 200);
