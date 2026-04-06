<?php
// /var/www/html/_api/imgcv_api.php
header('Content-Type: application/json; charset=UTF-8');

function jexit($a,$code=200){
  http_response_code($code);
  echo json_encode($a, JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES);
  exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') jexit(['ok'=>false,'error'=>'POST_ONLY'],405);

// secret check（ssh_execute.php と同じ方式）
$server_secret = @file_get_contents('/etc/niyantarose/ssh_secret.txt');
$server_secret = $server_secret ? trim($server_secret) : '';
$client_secret = $_SERVER['HTTP_X_VPS_SECRET'] ?? '';
$client_secret = is_string($client_secret) ? trim($client_secret) : '';

if ($server_secret === '' || $client_secret === '' || !hash_equals($server_secret, $client_secret)) {
  jexit(['ok'=>false,'error'=>'FORBIDDEN'],403);
}

// JSON body
$raw = file_get_contents('php://input');
$req = json_decode($raw, true);
if (!is_array($req)) jexit(['ok'=>false,'error'=>'BAD_JSON'],400);

$seller_id = trim((string)($req['seller_id'] ?? ''));
$items = $req['items'] ?? null;
$force = !empty($req['force']);
if ($seller_id === '' || !is_array($items) || !$items) jexit(['ok'=>false,'error'=>'seller_id/items required'],400);

// downloader
$BIN  = '/var/www/html/img/_bin/download_image_batch.sh';
if (!is_file($BIN)) jexit(['ok'=>false,'error'=>'missing downloader'],500);

// rclone env (www-data can read /etc/rclone/rclone.conf)
putenv('RCLONE_CONFIG=/etc/rclone/rclone.conf');

// remote name (default gdrive_ro:)
$remote = trim((string)($req['rclone_remote'] ?? 'gdrive_ro:'));
if ($remote === '') $remote = 'gdrive_ro:';

// index file (ID -> path)
$index = '/var/www/html/img/_cache/gdrive_index_ip.tsv';
if (!is_file($index)) {
  jexit(['ok'=>false,'error'=>'missing_index','index'=>$index],500);
}

$work = '/tmp/imgcv_api_' . getmypid() . '_' . time();
@mkdir($work, 0775, true);

$tsv = $work . '/list.tsv';
$fp = fopen($tsv, 'wb');
if (!$fp) jexit(['ok'=>false,'error'=>'cannot write tsv'],500);

$prep = [];

// helper: safe code/name
function safe_token($s, $re, $max=128){
  $s = trim((string)$s);
  if ($s === '' || strlen($s) > $max) return '';
  if (!preg_match($re, $s)) return '';
  return $s;
}

foreach ($items as $it) {
  $code = safe_token($it['code'] ?? '', '/^[A-Za-z0-9._-]{1,80}$/', 80);
  $name = safe_token($it['name'] ?? '', '/^[A-Za-z0-9._-]{1,120}$/', 120);
  $src  = trim((string)($it['src'] ?? ''));

  if ($code==='' || $name==='' || $src==='') {
    $prep[] = ['code'=>$code,'name'=>$name,'ok'=>false,'stage'=>'validate','err'=>'bad_fields'];
    continue;
  }

  // driveId:xxxx -> index -> rclone cat remote:path -> local file
  if (stripos($src, 'driveId:') === 0) {
    $id = substr($src, strlen('driveId:'));
    $id_esc = preg_replace('/[^A-Za-z0-9_-]/', '', (string)$id);
    if ($id_esc === '') {
      $prep[] = ['code'=>$code,'name'=>$name,'ok'=>false,'stage'=>'rclone','err'=>'bad_id'];
      continue;
    }

    // find path by grep "ID;"
    $cmdFind = "grep -m 1 -F " . escapeshellarg($id_esc . ";") . " " . escapeshellarg($index) . " 2>/dev/null";
    $line = trim((string)shell_exec($cmdFind));
    $path = '';
    if ($line !== '') {
      $parts = explode(';', $line, 2);
      if (count($parts) === 2) $path = trim($parts[1]);
    }

    if ($path === '') {
      $prep[] = ['code'=>$code,'name'=>$name,'ok'=>false,'stage'=>'rclone','err'=>'id_not_in_index','id'=>$id_esc];
      continue;
    }

    $path = ltrim($path, "/"); // safety
    $local = $work . '/src_' . preg_replace('/[^A-Za-z0-9_-]+/','_', $code.'_'.$name) . '.bin';

    $cmd = 'rclone cat ' . escapeshellarg($remote . $path)
         . ' > ' . escapeshellarg($local) . ' 2>/dev/null';
    $ret = 0;
    system($cmd, $ret);

    if ($ret !== 0 || !is_file($local) || filesize($local) < 1024) {
      $prep[] = ['code'=>$code,'name'=>$name,'ok'=>false,'stage'=>'rclone','err'=>'rclone_cat_failed','path'=>$path];
      continue;
    }

    $url = 'file://' . $local;
    fwrite($fp, $url . "\t" . $code . "\t" . $name . "\n");
    $prep[] = ['code'=>$code,'name'=>$name,'ok'=>true,'stage'=>'rclone','local'=>$local,'path'=>$path];
    continue;
  }

  // normal URL
  fwrite($fp, $src . "\t" . $code . "\t" . $name . "\n");
  $prep[] = ['code'=>$code,'name'=>$name,'ok'=>true,'stage'=>'url'];
}
fclose($fp);

// run downloader
$cmd = escapeshellcmd($BIN) . ' ' . ($force ? '--force ' : '') . escapeshellarg($tsv) . ' 2>&1';
$out = [];
$ret = 0;
exec($cmd, $out, $ret);

// parse results
$results = [];
foreach ($out as $line) {
  if (strpos($line, "IMGCV_RESULT\t") === 0) {
    $p = explode("\t", $line, 5);
    $results[] = [
      'code'=>$p[1] ?? '',
      'name'=>$p[2] ?? '',
      'status'=>$p[3] ?? '',
      'detail'=>$p[4] ?? ''
    ];
  }
}

// overall flags
$allPrepOk = true;
foreach ($prep as $x) { if (empty($x['ok'])) { $allPrepOk = false; break; } }

jexit([
  'ok'=>true,
  'seller_id'=>$seller_id,
  'force'=>$force,
  'rclone_remote'=>$remote,
  'all_prep_ok'=>$allPrepOk,
  'prep'=>$prep,
  'results'=>$results,
  'raw_tail'=>array_slice($out, -10),
  'workdir'=>$work
]);
