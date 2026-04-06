<?php
header('Content-Type: text/plain; charset=utf-8');

function respond($code, $msg) {
  http_response_code($code);
  echo $msg;
  exit;
}

// ====== secret check (header優先、POSTフォールバック) ======
$secret = $_SERVER['HTTP_X_VPS_SECRET'] ?? ($_POST['secret'] ?? '');
$secret = trim((string)$secret);

$EXPECTED = 'bce535b6993c5eda1ddf8eb9ff454dcc41ee142b354571df32b5f62515487dbf'; // ★ここに bce535... を入れる

if ($secret === '' || !hash_equals($EXPECTED, $secret)) {
  respond(403, "FORBIDDEN\n");
}

// ====== payload (base64 TSV) ======
$tsv_b64 = $_POST['tsv_b64'] ?? '';
$tsv_b64 = trim((string)$tsv_b64);
if ($tsv_b64 === '') respond(400, "NO_TSV\n");

// URL-safe base64も許可（- _）
$tsv_b64 = strtr($tsv_b64, '-_', '+/');

$tsv = base64_decode($tsv_b64, true);
if ($tsv === false) respond(400, "BAD_B64\n");

// サイズ制限（念のため）
if (strlen($tsv) > 1024 * 1024 * 2) respond(413, "TOO_LARGE\n"); // 2MB

// 改行などはTSVとして許可。ただしNULは拒否
if (strpos($tsv, "\0") !== false) respond(400, "REJECTED_NUL\n");

// ====== write tmp file ======
$tmp = '/tmp/imgcv_' . time() . '_' . bin2hex(random_bytes(4)) . '.tsv';
file_put_contents($tmp, $tsv);

// ====== run fixed script ======
$script = '/var/www/html/img/_bin/download_image_batch.sh';
$force = !empty($_POST['force']) ? '--force' : '';

$cmd = 'bash ' . escapeshellarg($script) . ' ' . $force . ' ' . escapeshellarg($tmp) . ' 2>&1';
$out = [];
$rc = 0;
exec($cmd, $out, $rc);

// cleanup
@unlink($tmp);

// ====== output ======
echo implode("\n", $out) . "\n";
if ($rc !== 0) {
  // 失敗でもログは返す（GAS側で判断）
  // http 200 のままにしておくと parse できる
}
