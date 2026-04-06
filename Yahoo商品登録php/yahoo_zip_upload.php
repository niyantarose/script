<?php

require_once __DIR__ . '/_vps_secret.php';
require_secret_or_403();

// ↓ここに追加
error_log('POST: ' . json_encode($_POST));
error_log('GET: ' . json_encode($_GET));

/**
 * yahoo_zip_upload.php (FULL FIXED)
 * - legacy: ZIP (upload/zip_url/zip_path) -> preprocess -> delete -> uploadPack -> submit
 * - action=submit: ZIPレス submit（単体/複数対応）
 *
 * 必須:
 * - X-VPS-Secret ヘッダが /etc/niyantarose/ssh_secret.txt と一致
 * - OAuth: access_token を POST で渡すか、ENVの refresh_token で自動取得
 */

ini_set('display_errors', 0);
ini_set('display_startup_errors', 0);
error_reporting(E_ALL);
header('Content-Type: application/json; charset=utf-8');
date_default_timezone_set('Asia/Tokyo');
@set_time_limit(0);
@ini_set('memory_limit', '1024M');

// =========================
// Constants
// =========================

const EP_UPLOAD_PACK   = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImagePack';
const EP_ITEMIMAGELIST = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/itemImageList';
const EP_DELETE_IMAGE  = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage';
const EP_SUBMIT_ITEM   = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem';
const EP_RESERVE_PUBLISH = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/reservePublish'; // ←これを追加！
const EP_TOKEN = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';
const ZIP_LIMIT_BYTES = 25 * 1024 * 1024; // 25MB
const RUN_BASE_DIR    = '/tmp/yahoo_zip_runs';
const TOKEN_CACHE     = '/tmp/yahoo_access_token_cache.json';

// SSRF 対策: zip_url は自ドメイン /dl/ だけ許可
const ALLOW_ZIP_URL_HOST   = 'img.niyantarose.com';
const ALLOW_ZIP_URL_PREFIX = '/dl/';

// zip_path 許可ディレクトリ
const ALLOW_ZIP_PATH_DIRS = [
  '/var/www/html/dl',
  '/var/www/private/temp_work',
];

// preprocess script
const PREPROCESS_BIN = '/usr/local/bin/yahoo_img_preprocess_dir.sh';

// =========================
// JSON responder
// =========================

function respond($arr, $code = 200) {
  http_response_code($code);
  echo json_encode($arr, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
  exit;
}

function now_ms() { return (int)(microtime(true) * 1000); }
function safe_mkdir($dir) { if (!is_dir($dir)) @mkdir($dir, 0777, true); }

// =========================
// Fatal handler: 致命エラーだけ JSON
// =========================
register_shutdown_function(function () {
  $e = error_get_last();
  if (!$e) return;

  $fatalTypes = [E_ERROR, E_PARSE, E_CORE_ERROR, E_COMPILE_ERROR, E_USER_ERROR];
  if (!in_array($e['type'], $fatalTypes, true)) return;
  http_response_code(500);
  echo json_encode([
    'ok' => false,
    'error' => 'FATAL',
    'fatal' => $e,
  ], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
});

// =========================
// Submit code parsing (submit専用)
// =========================
function parse_submit_item_codes(array $post): array {
  if (isset($post['item_codes']) && is_array($post['item_codes'])) {
    $arr = $post['item_codes'];
  } else {
    $raw = $post['item_codes'] ?? '';
    $raw = is_string($raw) ? trim($raw) : '';

    if ($raw === '') return [];

    if ($raw !== '' && ($raw[0] === '[' || $raw[0] === '{')) {
      $j = json_decode($raw, true);
      if (is_array($j)) {
        $arr = $j['item_codes'] ?? $j; 
      } else {
        $arr = [];
      }
    } else {
      $arr = preg_split('/[\s,;]+/u', $raw) ?: [];
    }
  }

  $out = [];
  $seen = [];
  foreach ($arr as $x) {
    $c = is_string($x) ? trim($x) : '';
    if ($c === '') continue;
    if (!preg_match('/^[A-Za-z0-9._-]{1,80}$/', $c)) continue;

    if (!isset($seen[$c])) {
      $seen[$c] = 1;
      $out[] = $c;
    }
  }
  return $out;
}

function pick_single_item_code(array $post): string {
  $c = $post['item_code'] ?? '';
  return is_string($c) ? trim($c) : '';
}

// =========================
// HTTP helper
// =========================
function curl_req($method, $url, $token, $postFields = null, $isMultipart = false, $extraHeaders = []) {
  $ch = curl_init($url);
  curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
  curl_setopt($ch, CURLOPT_HEADER, true);
  curl_setopt($ch, CURLOPT_TIMEOUT, 180);

  $headers = array_merge([
    'Accept: application/xml',
  ], $extraHeaders);

  if ($token !== null && $token !== '') {
    $headers[] = 'Authorization: Bearer ' . $token;
  }

  if ($method === 'POST') {
    curl_setopt($ch, CURLOPT_POST, true);
    if ($postFields !== null) {
      curl_setopt($ch, CURLOPT_POSTFIELDS, $postFields);
    if (!$isMultipart) {
        $headers[] = 'Content-Type: application/x-www-form-urlencoded';
      }
    }
  } else {
    curl_setopt($ch, CURLOPT_HTTPGET, true);
  }

  curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
  $resp  = curl_exec($ch);
  $errno = curl_errno($ch);
  $err   = curl_error($ch);
  $info  = curl_getinfo($ch);

  curl_close($ch);
  if ($errno) {
    return [
      'ok' => false,
      'http' => 0,
      'errno' => $errno,
      'error' => $err,
      'body' => '',
      'raw' => '',
    ];
  }

  $headerSize = (int)($info['header_size'] ?? 0);
  $body = substr($resp ?: '', $headerSize);
  return [
    'ok' => true,
    'http' => (int)($info['http_code'] ?? 0),
    'body' => $body,
    'raw' => $resp,
  ];
}

// =========================
// URL helpers (reupload_urls)
// =========================
const REUPLOAD_MAX_TOTAL  = 21; 
const REUPLOAD_MAX_NOTICE = 5;  
const REUPLOAD_MAX_DETAIL = 20; 

const ALLOW_SRC_HOSTS = [
  'img.niyantarose.com',
  'item-shopping.c.yimg.jp',
  'drive.google.com',
];

function url_is_allowed($url) {
  $u = @parse_url($url);
  if (!$u || empty($u['scheme']) || empty($u['host'])) return false;
  $scheme = strtolower($u['scheme']);
  $host   = strtolower($u['host']);
  if ($scheme !== 'https') return false;
  return in_array($host, ALLOW_SRC_HOSTS, true);
}

function split_urls_any($raw): array {
  if (is_array($raw)) {
    $arr = $raw;
  } else {
    $s = is_string($raw) ? $raw : '';
    $s = str_replace(["\r", "\t"], ["\n", " "], $s);
    $arr = preg_split('/[;\n\s]+/u', $s) ?: [];
  }
  $out = [];
  $seen = [];

  foreach ($arr as $x) {
    $u = is_string($x) ? trim($x) : '';
    if ($u === '') continue;
    if (!preg_match('~^https?://~i', $u)) continue;
    if (!isset($seen[$u])) { $seen[$u] = 1; $out[] = $u; }
  }
  return $out;
}

function is_notice_url($url): bool {
  $u = @parse_url($url);
  $path = $u['path'] ?? '';
  $base = strtolower(basename($path));
  return (strpos($base, 'notice_') === 0) || (strpos($base, '免責') !== false);
}

function compute_notice_nums($count): array {
  $count = max(0, min(REUPLOAD_MAX_NOTICE, (int)$count));
  if ($count <= 0) return [];
  $start = 100 - $count;
  $nums = [];
  for ($i=0; $i<$count; $i++) $nums[] = $start + $i;
  return $nums;
}


// ★ここに追加
function google_drive_resolve_url_($url) {
  $ch = curl_init($url);
  curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
  curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
  curl_setopt($ch, CURLOPT_TIMEOUT, 30);
  curl_setopt($ch, CURLOPT_COOKIEJAR, '/tmp/gdrive_cookie.txt');
  curl_setopt($ch, CURLOPT_COOKIEFILE, '/tmp/gdrive_cookie.txt');
  $body = curl_exec($ch);
  curl_close($ch);

  if (preg_match('/confirm=([0-9A-Za-z_\-]+)/', $body, $m)) {
    $sep = strpos($url, '?') !== false ? '&' : '?';
    return $url . $sep . 'confirm=' . $m[1];
  }

  return $url;
}


function download_to_file($url, $dst, &$info = []) {
  $info = ['url'=>$url, 'dst'=>$dst, 'ok'=>false, 'http'=>0, 'err'=>''];
  if (!url_is_allowed($url)) {
    $info['err'] = 'url not allowed';
    return false;
  }

  // ★ Google Drive の確認トークン対応
  if (strpos($url, 'drive.google.com') !== false) {
    $url = google_drive_resolve_url_($url);
  }


  $fp = @fopen($dst, 'wb');
  if (!$fp) {
    $info['err'] = 'fopen failed';
    return false;
  }

  $ch = curl_init($url);
  curl_setopt($ch, CURLOPT_FILE, $fp);
  curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
  curl_setopt($ch, CURLOPT_TIMEOUT, 180);
  curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, 20);
  curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, true);
  curl_setopt($ch, CURLOPT_FAILONERROR, false);
  curl_setopt($ch, CURLOPT_USERAGENT, 'niyantarose-reupload/1.0');

  $ok = curl_exec($ch);
  $http = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
  $err  = curl_error($ch);
  curl_close($ch);
  fclose($fp);

  $info['http'] = $http;
  if (!$ok || $http < 200 || $http >= 300) {
    @unlink($dst);
    $info['err'] = $err ?: ("http=" . $http);
    return false;
  }

  clearstatcache(true, $dst);
  $sz = @filesize($dst);
  if ($sz === false || $sz <= 0) {
    @unlink($dst);
    $info['err'] = 'downloaded size=0';
    return false;
  }

  $info['ok'] = true;
  $info['size'] = $sz;
  return true;
}

function parse_reupload_items(array $post): array {
  $raw = $post['items_json'] ?? '';
  $raw = is_string($raw) ? trim($raw) : '';
  if ($raw === '') return [];

  $j = json_decode($raw, true);
  if (!is_array($j)) return [];

  $out = [];
  foreach ($j as $it) {
    if (!is_array($it)) continue;
    $code = trim((string)($it['item_code'] ?? ''));
    if ($code === '') continue;
    if (!preg_match('/^[A-Za-z0-9._-]{1,80}$/', $code)) continue;

    $s = $it['s'] ?? ($it['S'] ?? '');
    $t = $it['t'] ?? ($it['T'] ?? '');

    $sUrls = split_urls_any($s);
    $tUrls = split_urls_any($t);

    $out[] = [
      'item_code' => $code,
      's' => $sUrls ? $sUrls[0] : '',
      't' => $tUrls,
    ];
  }
  return $out;
}

function build_workdir_from_items(array $items, string $workDir, array &$downloadSummary) {
  $downloadSummary = [];
  $itemCodes = [];

  foreach ($items as $it) {
    $code = $it['item_code'];
    $codeLower = strtolower($code);
    $itemCodes[] = $code;

    $s = $it['s'] ?? '';
    $t = $it['t'] ?? [];
    $t = is_array($t) ? $t : split_urls_any($t);

    $prod = [];
    $notice = [];
    foreach ($t as $u) {
      if (is_notice_url($u)) $notice[] = $u;
      else $prod[] = $u;
    }

    if ($s === '' && $prod) {
      $s = array_shift($prod);
    }

    if ($s !== '') {
      $dst = $workDir . '/' . $codeLower . '.jpg';
      $info = [];
      $ok = download_to_file($s, $dst, $info);
      $downloadSummary[] = array_merge(['item_code'=>$code, 'slot'=>'S', 'name'=>basename($dst)], $info);
    }

    $prod = array_slice($prod, 0, REUPLOAD_MAX_DETAIL);

    for ($i=0; $i<count($prod); $i++) {
      $num = $i + 1;
      $dst = $workDir . '/' . $codeLower . '_' . $num . '.jpg';
      $info = [];
      $ok = download_to_file($prod[$i], $dst, $info);
      $downloadSummary[] = array_merge(['item_code'=>$code, 'slot'=>'T', 'name'=>basename($dst)], $info);
    }

    $notice = array_slice($notice, 0, REUPLOAD_MAX_NOTICE);
    $nums = compute_notice_nums(count($notice));

    for ($k=0; $k<count($notice); $k++) {
      $num = $nums[$k];
      $dst = $workDir . '/' . $codeLower . '_' . $num . '.jpg';
      $info = [];
      $ok = download_to_file($notice[$k], $dst, $info);
      $downloadSummary[] = array_merge(['item_code'=>$code, 'slot'=>'NOTICE', 'name'=>basename($dst)], $info);
    }
  }

  return $itemCodes;
}

function yahoo_xml_ok($xml): bool {
  if (!is_string($xml) || $xml === '') return false;
  if (stripos($xml, '<Error>') !== false) return false;
  if (stripos($xml, '<Code>') !== false && stripos($xml, '<Message>') !== false) return false;
  return true;
}

function yahoo_error_code($xmlStr): string {
  if (!is_string($xmlStr) || $xmlStr === '') return '';
  if (preg_match('~<Code>\s*([^<\s]+)\s*</Code>~i', $xmlStr, $m)) {
    return trim($m[1]);
  }
  return '';
}

function xml_load($xmlStr) {
  libxml_use_internal_errors(true);
  $xml = @simplexml_load_string($xmlStr);
  return $xml ?: null;
}

function chunk_array($arr, $size) {
  $out = [];
  $buf = [];
  foreach ($arr as $v) {
    $buf[] = $v;
    if (count($buf) >= $size) { $out[] = $buf; $buf = []; }
  }
  if ($buf) $out[] = $buf;
  return $out;
}

// =========================
// OAuth token refresh (ENV + cache)
// =========================
function load_cached_token() {
  if (!file_exists(TOKEN_CACHE)) return null;
  $j = json_decode(@file_get_contents(TOKEN_CACHE), true);
  if (!is_array($j)) return null;

  $token = $j['access_token'] ?? '';
  $expAt = (int)($j['expires_at'] ?? 0);

  if ($token !== '' && time() < ($expAt - 60)) return $token;
  return null;
}

function save_cached_token($token, $expiresIn) {
  $expiresAt = time() + (int)$expiresIn;
  @file_put_contents(TOKEN_CACHE, json_encode([
    'access_token' => $token,
    'expires_at' => $expiresAt,
    'saved_at' => time(),
  ], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES));
}

function refresh_access_token_from_env(&$debug = []) {
  $cached = load_cached_token();
  if ($cached) {
    $debug['token_source'] = 'cache';
    return $cached;
  }

  $cid = getenv('YAHOO_CLIENT_ID') ?: '';
  $sec = getenv('YAHOO_CLIENT_SECRET') ?: '';
  $rt  = getenv('YAHOO_REFRESH_TOKEN') ?: '';

  if ($cid === '' || $sec === '' || $rt === '') {
    $debug['token_source'] = 'none';
    $debug['env_missing'] = [
      'YAHOO_CLIENT_ID' => ($cid === ''),
      'YAHOO_CLIENT_SECRET' => ($sec === ''),
      'YAHOO_REFRESH_TOKEN' => ($rt === ''),
    ];
    return '';
  }

  $post = http_build_query([
    'grant_type' => 'refresh_token',
    'refresh_token' => $rt,
  ]);

  $basic = base64_encode($cid . ':' . $sec);
  $r = curl_req('POST', EP_TOKEN, null, $post, false, [
    'Content-Type: application/x-www-form-urlencoded',
    'Authorization: Basic ' . $basic,
    'Accept: application/json',
  ]);

  $debug['token_http'] = $r['http'];

  if (!$r['ok'] || $r['http'] !== 200) {
    $debug['token_error_body_tail'] = substr($r['body'], -800);
    return '';
  }

  $j = json_decode($r['body'], true);
  $token = $j['access_token'] ?? '';
  $expiresIn = $j['expires_in'] ?? 0;

  if ($token !== '' && $expiresIn) {
    save_cached_token($token, $expiresIn);
    $debug['token_source'] = 'refresh';
    return $token;
  }

  $debug['token_source'] = 'refresh_parse_failed';
  $debug['token_error_body_tail'] = substr($r['body'], -800);
  return '';
}

// =========================
// ZIP source resolve (upload / zip_url / zip_path)
// =========================
function realpath_in_allowed_dirs($path, $allowedDirs) {
  $rp = realpath($path);
  if (!$rp) return [null, 'realpath failed'];

  foreach ($allowedDirs as $d) {
    $rd = realpath($d);
    if (!$rd) continue;
    $prefix = rtrim($rd, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR;
    if (strpos($rp, $prefix) === 0) return [$rp, null];
  }
  return [null, 'zip_path not allowed'];
}

function resolve_zip_source(&$meta) {
  $meta = [
    'source' => null,
    'zip_url' => null,
    'zip_path' => null,
    'upload' => false,
  ];

  if (isset($_FILES['zip']) && is_array($_FILES['zip']) && is_uploaded_file($_FILES['zip']['tmp_name'] ?? '')) {
    $meta['source'] = 'upload';
    $meta['upload'] = true;
    return [$_FILES['zip']['tmp_name'], null];
  }

  $zipUrl = (string)($_POST['zip_url'] ?? '');
  if ($zipUrl !== '') {
    $u = parse_url($zipUrl);
    if (!$u || empty($u['scheme']) || empty($u['host']) || empty($u['path'])) {
      return [null, 'invalid zip_url'];
    }
    $scheme = strtolower($u['scheme']);
    $host   = strtolower($u['host']);
    $path   = $u['path'];

    if ($scheme !== 'https' || $host !== ALLOW_ZIP_URL_HOST || strpos($path, ALLOW_ZIP_URL_PREFIX) !== 0) {
      return [null, 'zip_url not allowed'];
    }

    $base = basename($path);
    if (!preg_match('/\.zip$/i', $base)) {
      return [null, 'zip_url must end with .zip'];
    }

    $local = '/var/www/html/dl/' . $base;
    clearstatcache(true, $local);
    if (!is_file($local)) {
      return [null, 'zip file not found'];
    }

    $meta['source'] = 'zip_url';
    $meta['zip_url'] = $zipUrl;
    $meta['zip_path'] = $local;
    return [$local, null];
  }

  $zipPath = (string)($_POST['zip_path'] ?? '');
  if ($zipPath !== '') {
    [$rp, $err] = realpath_in_allowed_dirs($zipPath, ALLOW_ZIP_PATH_DIRS);
    if ($err) return [null, $err];
    clearstatcache(true, $rp);
    if (!is_file($rp)) return [null, 'zip file not found'];

    $meta['source'] = 'zip_path';
    $meta['zip_path'] = $rp;
    return [$rp, null];
  }

if ($action === 'delete' || $action === 'submit') {
  return [null, null];
}

  return [null, 'zip not provided'];
}


// =========================
// Parse item codes from zip (legacy)
// =========================
function parse_item_codes_from_zip($zipPath) {
  $za = new ZipArchive();
  if ($za->open($zipPath) !== true) return [[], 'ZIP_OPEN_FAIL'];

  $codes = [];
  for ($i = 0; $i < $za->numFiles; $i++) {
    $stat = $za->statIndex($i);
    if (!$stat || empty($stat['name'])) continue;
    $name = $stat['name'];

    if (substr($name, -1) === '/' || str_starts_with($name, '__MACOSX/')) continue;

    $base = basename($name);
    if (!preg_match('/\.(jpe?g|png|gif|webp)$/i', $base)) continue;

    $stem = preg_replace('/\.(jpe?g|png|gif|webp)$/i', '', $base);

    if (preg_match('/^(.+?)_\d+$/', $stem, $m)) $code = $m[1];
    else $code = $stem;

    $code = trim($code);
    if (stripos($code, 'notice_') === 0) continue;

    if ($code !== '') $codes[$code] = true;
  }
  $za->close();
  return [array_keys($codes), null];
}

function inject_notice_images_into_workdir(array $itemCodes, string $workDir, bool $strict, array &$noticeSummary) {
  $PRODUCT_BASE = '/data/yahoo_images';
  $NOTICE_BASE  = '/data/yahoo_images/notice_sets';

  foreach ($itemCodes as $code) {
    $code = trim($code);
    if ($code === '') continue;

    $codeDir = $PRODUCT_BASE . '/' . $code;
    $idFile  = $codeDir . '/notice_id.txt';
    if (!is_file($idFile)) continue;

    $noticeId = trim(@file_get_contents($idFile));
    if ($noticeId === '' || strtolower($noticeId) === 'none') continue;

    $setDir = $NOTICE_BASE . '/' . $noticeId;
    $files = glob($setDir . '/*.jpg');
    sort($files, SORT_STRING);

    if (!$files) {
      $noticeSummary[] = ['item_code'=>$code, 'notice_id'=>$noticeId, 'ok'=>false, 'reason'=>'notice set empty'];
      if ($strict) throw new Exception("notice set empty: {$noticeId} for {$code}");
      continue;
    }

    $max = min(5, count($files));
    $startNum = 100 - $max;

    $codeLower = strtolower($code);
    for ($i=0; $i<$max; $i++) {
      $num = $startNum + $i;
      $dst = $workDir . '/' . $codeLower . '_' . $num . '.jpg';
      @unlink($dst);
      if (!copy($files[$i], $dst)) {
        $noticeSummary[] = ['item_code'=>$code, 'notice_id'=>$noticeId, 'ok'=>false, 'reason'=>'copy failed', 'src'=>$files[$i], 'dst'=>$dst];
        if ($strict) throw new Exception("notice copy failed: {$code} {$noticeId}");
        continue 2;
      }
    }

    $noticeSummary[] = ['item_code'=>$code, 'notice_id'=>$noticeId, 'ok'=>true, 'count'=>$max, 'nums'=>[$startNum, 99]];
  }
}

// =========================
// submit one item (ZIPレス用)
// =========================
function submit_one_item($sellerId, $itemCode, $token) {
  $post = http_build_query([
    'seller_id' => $sellerId,
    'item_code' => $itemCode,
  ]);
  $r = curl_req('POST', EP_SUBMIT_ITEM, $token, $post, false);
  $ok = ($r['ok'] && $r['http'] === 200 && yahoo_xml_ok($r['body']));
  return [
    'item_code' => $itemCode,
    'http' => $r['http'],
    'ok' => $ok,
    'body_tail' => substr($r['body'], -600),
    'curl_error' => $r['ok'] ? '' : ($r['error'] ?? ''),
  ];
}

// =========================
// DELETE helper: 指定item_codesの画像を全削除（legacy / delete共用）
// =========================
function delete_images_for_codes(array $itemCodes, string $sellerId, string $token): array {
  $deleteSummary = [];
  $totalDeleteTargets = 0;
  $deleteAllOk = true;

  foreach ($itemCodes as $code) {
    $codeLower = strtolower($code);
    $itemDeleteOk = true;

    $url = EP_ITEMIMAGELIST . '?' . http_build_query([
      'seller_id' => $sellerId,
      'query'     => $codeLower,
      'results'   => 100,
      'start'     => 1,
    ]);

    $r = curl_req('GET', $url, $token, null, false);
    $targets = [];

    if (!($r['ok'] && $r['http'] === 200)) {
      $itemDeleteOk = false;
    } else {
      $xml = xml_load($r['body']);
      if (!$xml) {
        $itemDeleteOk = false;
      } else {
        $results = $xml->xpath('//ResultSet/Result') ?: [];
        foreach ($results as $res) {
          $id   = (string)($res->Id ?? '');
          $name = (string)($res->Name ?? '');
          if ($id === '' || $name === '') continue;

          if (preg_match('/^' . preg_quote($codeLower, '/') . '(\.|_)/', strtolower($name))) {
            $targets[$id] = true;
          }
        }
      }
    }

    $ids = array_keys($targets);
    $totalDeleteTargets += count($ids);

    $delChunks = chunk_array($ids, 100);
    $delResults = [];

    foreach ($delChunks as $chunk) {
      if (!$chunk) continue;

      $post = http_build_query([
        'seller_id' => $sellerId,
        'image_id'  => implode(',', $chunk),
      ]);

      $dr = curl_req('POST', EP_DELETE_IMAGE, $token, $post, false);
      $chunkOk = ($dr['ok'] && $dr['http'] === 200 && yahoo_xml_ok($dr['body']));
      if (!$chunkOk) $itemDeleteOk = false;

      $delResults[] = [
        'http' => $dr['http'],
        'ok'   => $chunkOk,
        'body_tail' => substr($dr['body'], -600),
      ];
    }

    if (!$itemDeleteOk) $deleteAllOk = false;

    $deleteSummary[] = [
      'item_code' => $code,
      'ok'        => $itemDeleteOk,
      'found'     => count($ids),
      'list_http' => $r['http'],
      'list_body_tail' => substr($r['body'], -600),
      'delete_calls' => $delResults,
    ];
  }

  return [
    'ok' => $deleteAllOk,
    'targets_total' => $totalDeleteTargets,
    'details' => $deleteSummary,
  ];
}


// =========================
// Main
// =========================
if (!defined('IS_INCLUDED')) { // ←★ここに追加！（ここから下を全て囲む）

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
  respond(['ok' => false, 'error' => 'Method not allowed'], 405);
}

// --- 認証 ---
require_secret_or_403();

// --- ▼ 修正箇所：共通パラメータとトークンを最初に全て取得する ▼ ---

// URL・POST両方から受け取れるよう $_REQUEST を使用
$action = $_REQUEST['action'] ?? 'legacy';
$action = is_string($action) ? trim($action) : 'legacy';

$sellerId = trim((string)($_REQUEST['seller_id'] ?? ''));

// access_tokenの取得と自動更新
$token = trim((string)($_REQUEST['access_token'] ?? ''));
$debugTok = [];
if ($token === '') {
  $token = refresh_access_token_from_env($debugTok);
  if ($token === '') {
    respond([
      'ok' => false,
      'error' => 'access_token missing and refresh failed',
      'token_debug' => $debugTok,
    ], 401);
  }
}

// action が 'pipeline' 以外で、seller_id が無ければここで即座にエラー
if ($sellerId === '') {
    respond(['ok' => false, 'error' => 'seller_id required'], 400);
}

// --- ▲ 修正箇所おわり ▲ ---


// =====================================================
// action=pipeline : VPSが Drive(rclone/SA) -> preprocess -> zip -> uploadPack -> (optional)submit
// =====================================================
if ($action === 'pipeline') {

  $expected = trim(getenv('IMGCV_VPS_SECRET') ?: '');
  if ($expected !== '') {
    $got = $_SERVER['HTTP_X_VPS_SECRET'] ?? '';
    if (!hash_equals($expected, $got)) {
      respond(['ok' => false, 'error' => 'FORBIDDEN'], 403);
    }
  }

  $itemsRaw = (string)($_POST['items'] ?? '');
  $items = json_decode($itemsRaw, true);
  if (!is_array($items) || count($items) === 0) {
    respond(['ok' => false, 'error' => 'items required', 'v' => 'pipeline-missing-items'], 400);
  }

  $skipSubmit = !empty($_REQUEST['skip_submit']);
  $skipDelete = !empty($_REQUEST['skip_delete']);
  $strict = !empty($_REQUEST['strict']);

  $pipeline = getenv('IMGCV_PIPELINE') ?: '/usr/local/bin/yahoo_img_pipeline.sh';
  if (!is_file($pipeline)) {
    respond(['ok' => false, 'error' => 'pipeline not found', 'path' => $pipeline, 'v' => 'pipeline-missing-script'], 500);
  }

  // items -> TSV(TYPE<TAB>SRC<TAB>CODE<TAB>NAME)
  $rows = [];
  $codesMap = [];

  foreach ($items as $it) {
    if (!is_array($it)) continue;

    $code = trim((string)($it['code'] ?? $it['item_code'] ?? ''));
    $src  = trim((string)($it['src'] ?? $it['url'] ?? ''));
    $name = trim((string)($it['name'] ?? ''));
    $type = strtolower(trim((string)($it['type'] ?? '')));

    if ($code === '' || $src === '') continue;
    if (!preg_match('/^[A-Za-z0-9._-]{1,80}$/', $code)) continue;

    if ($type === '') {
      if (str_starts_with($src, 'driveId:')) {
        $type = 'id';
        $src = substr($src, 8);
      } elseif (preg_match('~^https?://~i', $src)) {
        $type = 'url';
      } elseif (preg_match('/^[A-Za-z0-9_-]{10,}$/', $src)) {
        $type = 'id';
      } else {
        continue;
      }
    } else {
      if (!in_array($type, ['id', 'url', 'rcpath'], true)) continue;
      if ($type === 'id' && str_starts_with($src, 'driveId:')) {
        $src = substr($src, 8);
      }
    }

    if ($name === '') $name = $code . '.jpg';
    $name = basename($name);
    if ($name === '') $name = $code . '.jpg';

    $rows[] = implode("\t", [$type, $src, $code, $name]);
    $codesMap[$code] = 1;
  }

  if (count($rows) === 0) {
    respond(['ok' => false, 'error' => 'no valid items(code/src)', 'v' => 'pipeline-invalid-items'], 400);
  }

  $itemCodes = array_keys($codesMap);

  $runId = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
  $workDir = '/var/www/private/temp_work/pipeline_' . $runId;
  safe_mkdir($workDir);

  $tmpTsv = tempnam(sys_get_temp_dir(), 'imgcv_tsv_');
  file_put_contents($tmpTsv, implode("\n", $rows) . "\n");

  $cmd = escapeshellcmd($pipeline)
       . ' ' . escapeshellarg($tmpTsv)
       . ' ' . escapeshellarg($workDir);

  $out = [];
  $rc = 0;
  exec($cmd . ' 2>&1', $out, $rc);
  @unlink($tmpTsv);

  $stdout = trim(implode("\n", $out));
  if ($rc !== 0) {
    respond([
      'ok' => false,
      'error' => 'pipeline failed',
      'rc' => $rc,
      'log' => mb_substr($stdout, 0, 4000),
      'v' => 'pipeline-rc-nonzero'
    ], 500);
  }

  $lines = ($stdout === '') ? [] : preg_split('/\r?\n/u', $stdout);
  $dlOk = 0;
  $dlNg = 0;
  foreach ($lines as $ln) {
    $ln = trim((string)$ln);
    if ($ln === '') continue;
    if (preg_match('/^OK[\t ]+/u', $ln)) $dlOk++;
    if (preg_match('/^NG[\t ]+/u', $ln)) $dlNg++;
  }

  $imgRoot = $workDir . '/img';
  $zipOut = $workDir . '/pipeline_upload.zip';
  $zb = new ZipArchive();
  if ($zb->open($zipOut, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    respond(['ok' => false, 'error' => 'zip create failed', 'v' => 'pipeline-zip-open-fail'], 500);
  }

  $zipFiles = 0;
  if (is_dir($imgRoot)) {
    $it = new RecursiveIteratorIterator(
      new RecursiveDirectoryIterator($imgRoot, FilesystemIterator::SKIP_DOTS)
    );
    foreach ($it as $f) {
      if (!$f->isFile()) continue;
      $fn = $f->getFilename();
      if (!preg_match('/\.(jpe?g|png|webp|gif)$/i', $fn)) continue;
      $zb->addFile($f->getPathname(), $fn); 
      $zipFiles++;
    }
  }
  $zb->close();

  if ($zipFiles === 0) {
    respond([
      'ok' => false,
      'error' => 'no images from pipeline',
      'raw_tail' => array_slice($lines, -80),
      'v' => 'pipeline-empty'
    ], 500);
  }

$deleteResult = ['ok' => true, 'targets_total' => 0, 'details' => []];
/* 以下のifブロックを丸ごとコメントアウト 
if (!$skipDelete) {
  $deleteResult = delete_images_for_codes($itemCodes, $sellerId, $token);
}
*/


  $upUrl = EP_UPLOAD_PACK . '?seller_id=' . rawurlencode($sellerId);
  $cfile = new CURLFile($zipOut, 'application/zip', 'images.zip');
  $up = curl_req('POST', $upUrl, $token, ['file' => $cfile], true);
  $uploadOk = ($up['ok'] && $up['http'] === 200 && yahoo_xml_ok($up['body']));

  $submitSummary = [];
  $okCount = 0;
  $ngCount = 0;

  if (!$skipSubmit && $uploadOk) {
    foreach ($itemCodes as $code) {
      $d = submit_one_item($sellerId, $code, $token);
      $submitSummary[] = $d;
      if (!empty($d['ok'])) $okCount++;
      else $ngCount++;
    }
  }

  $allSubmitOk = ($ngCount === 0);
  $anySubmitOk = ($okCount > 0);

  $finalOk = $skipSubmit
    ? $uploadOk
    : ($strict ? ($uploadOk && $allSubmitOk) : ($uploadOk && $anySubmitOk));

  respond([
    'ok' => $finalOk,
    'action' => 'pipeline',
    'strict' => $strict,
    'skip_submit' => $skipSubmit,
    'skip_delete' => $skipDelete,
    'run_id' => $runId,
    'work_dir' => $workDir,
    'saved_zip' => $zipOut,
    'item_codes' => $itemCodes,
    'download' => ['ok' => $dlOk, 'ng' => $dlNg],
    'zip' => ['files' => $zipFiles],
    'delete' => $deleteResult,
    'upload_pack' => [
      'http' => $up['http'],
      'ok' => $uploadOk,
      'body_tail' => substr($up['body'], -800),
    ],
    'submit' => [
      'ok' => $skipSubmit ? null : $allSubmitOk,
      'skipped' => $skipSubmit,
      'ok_count' => $okCount,
      'ng_count' => $ngCount,
      'details' => $submitSummary,
    ],
    'token_debug' => $debugTok,
    'raw_tail' => array_slice($lines, -80),
    'v' => 'pipeline-v2'
  ], 200);
}


// =====================================================
// action=submit (ZIPレス) : item_code / item_codes をsubmit
// =====================================================
if ($action === 'submit') {
  $itemCodes = parse_submit_item_codes($_POST);

  if (!$itemCodes) {
    $single = pick_single_item_code($_POST);
    if ($single !== '') $itemCodes = [$single];
  }

  if (!$itemCodes) {
    respond([
      'ok' => false,
      'error' => 'item_code(s) required',
      'post_keys' => array_keys($_POST),
      'v' => 'submit-branch-missing',
    ], 400);
  }

  if (count($itemCodes) > 200) {
    respond(['ok'=>false,'error'=>'too_many_item_codes','count'=>count($itemCodes)], 400);
  }

  $strict = !empty($_REQUEST['strict']);

  $details = [];
  $okCount = 0;
  $ngCount = 0;

  foreach ($itemCodes as $code) {
    $d = submit_one_item($sellerId, $code, $token);
    $details[] = $d;
    if (!empty($d['ok'])) $okCount++;
    else $ngCount++;
  }

  $allOk   = ($ngCount === 0);
  $anyOk   = ($okCount > 0);
  $finalOk = $strict ? $allOk : $anyOk;

  respond([
    'ok' => $finalOk,
    'action' => 'submit',
    'submit' => [
      'count' => count($itemCodes),
      'ok_count' => $okCount,
      'ng_count' => $ngCount,
      'strict' => $strict,
      'details' => $details,
    ],
    'token_debug' => $debugTok,
    'v' => 'submit-strict-v1',
  ], 200);
}

// =====================================================
// action=delete (ZIPレス) : item_code / item_codes の画像を削除のみ
// =====================================================
if ($action === 'delete') {
  $itemCodes = parse_submit_item_codes($_POST);

  if (!$itemCodes) {
    $single = pick_single_item_code($_POST);
    if ($single !== '') $itemCodes = [$single];
  }

  if (!$itemCodes) {
    respond([
      'ok' => false,
      'error' => 'item_code(s) required for delete',
      'post_keys' => array_keys($_POST),
      'v' => 'delete-branch-missing',
    ], 400);
  }

  if (count($itemCodes) > 200) {
    respond(['ok'=>false,'error'=>'too_many_item_codes','count'=>count($itemCodes)], 400);
  }

$t0 = now_ms();
// $deleteResult = delete_images_for_codes($itemCodes, $sellerId, $token); // ←ここをコメントアウト！
$deleteResult = ['ok' => true, 'targets_total' => 0, 'details' => []]; // ←追記
  $elapsed = now_ms() - $t0;

  respond([
    'ok' => $deleteResult['ok'],
    'action' => 'delete',
    'item_codes' => $itemCodes,
    'delete' => $deleteResult,
    'elapsed_ms' => $elapsed,
    'token_debug' => $debugTok,
    'v' => 'delete-v1',
  ], 200);
}

// =====================================================
// action=reserve_publish (ZIPレス) : 全反映予約
// =====================================================
if ($action === 'reserve_publish') {
  $post = http_build_query([
    'seller_id' => $sellerId,
    'mode'      => 1,
  ]);

  $t0 = now_ms();
  $r = curl_req('POST', EP_RESERVE_PUBLISH, $token, $post, false);
  $elapsed = now_ms() - $t0;

  $ok = false;
  $reserveTime = '';
  if ($r['ok'] && $r['http'] === 200) {
    $xml = xml_load($r['body']);
    if ($xml && (string)$xml->Status === 'OK') {
      $ok = true;
      $reserveTime = (string)$xml->ReserveTime;
    }
  }

  respond([
    'ok' => $ok,
    'action' => 'reserve_publish',
    'reserve_time' => $reserveTime,
    'http' => $r['http'],
    'body_tail' => substr($r['body'], -600),
    'elapsed_ms' => $elapsed,
    'token_debug' => $debugTok,
    'v' => 'reserve-publish-v1',
  ], 200);
}

// =====================================================
// action=reupload_urls : URL群(S/T/免責) -> 全削除 -> ZIP化 -> uploadPack -> submit
// =====================================================
if ($action === 'reupload_urls') {
  $strict = !empty($_REQUEST['strict']);

  $items = parse_reupload_items($_POST);
  if (!$items) {
    respond([
      'ok' => false,
      'error' => 'items_json required (JSON array)',
      'example' => '[{"item_code":"TWF0001","s":"https://...","t":"https://...\\nhttps://...;https://..."}]',
      'post_keys' => array_keys($_POST),
      'v' => 'reupload_urls-missing-items',
    ], 400);
  }

  if (count($items) > 200) {
    respond(['ok'=>false,'error'=>'too_many_items','count'=>count($items)], 400);
  }

  safe_mkdir(RUN_BASE_DIR);
  $runId  = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
  $runDir = rtrim(RUN_BASE_DIR, '/') . '/' . $runId;
  safe_mkdir($runDir);

  $workDir = $runDir . '/work';
  safe_mkdir($workDir);

  $downloadSummary = [];
  $itemCodes = build_workdir_from_items($items, $workDir, $downloadSummary);
  $itemCodes = array_values(array_unique($itemCodes));

  $noticeSummary = [];
  try {
    inject_notice_images_into_workdir($itemCodes, $workDir, $strict, $noticeSummary);
  } catch (Throwable $e) {
    if ($strict) {
      respond([
        'ok' => false,
        'error' => 'NOTICE_INJECT_FAIL',
        'message' => $e->getMessage(),
        'notice' => $noticeSummary,
      ], 500);
    }
    $noticeSummary[] = [
      'item_code' => '(unknown)',
      'notice_id' => '',
      'ok' => false,
      'reason' => 'exception',
      'message' => $e->getMessage(),
    ];
  }

  if (!is_file(PREPROCESS_BIN)) {
    respond(['ok'=>false,'error'=>'missing preprocess bin','bin'=>PREPROCESS_BIN], 500);
  }
  $cmd = PREPROCESS_BIN . ' ' . escapeshellarg($workDir) . ' 2>&1';
  $ppOut = [];
  $ppRet = 0;
  exec($cmd, $ppOut, $ppRet);

  $zipOut = $runDir . '/input_processed.zip';
  $zb = new ZipArchive();
  if ($zb->open($zipOut, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    respond(['ok'=>false,'error'=>'ZIP_CREATE_FAIL_FOR_REUPLOAD'], 500);
  }

  $added = 0;
  $dh = opendir($workDir);
  if ($dh !== false) {
    while (($fn = readdir($dh)) !== false) {
      if ($fn === '.' || $fn === '..') continue;
      if (!preg_match('/\.(jpe?g)$/i', $fn)) continue;
      $full = $workDir . '/' . $fn;
      if (is_file($full)) { $zb->addFile($full, $fn); $added++; }
    }
    closedir($dh);
  }
  $zb->close();

  if ($ppRet !== 0 || $added === 0) {
    respond([
      'ok' => false,
      'error' => 'PREPROCESS_FAIL_OR_EMPTY',
      'preprocess_ret' => $ppRet,
      'added_jpg' => $added,
      'preprocess_tail' => array_slice($ppOut, -80),
      'download_head' => array_slice($downloadSummary, 0, 20),
    ], 500);
  }

  $t0 = now_ms();
  $deleteResult = delete_images_for_codes($itemCodes, $sellerId, $token);

  $upUrl = EP_UPLOAD_PACK . '?seller_id=' . rawurlencode($sellerId);
  $cfile = new CURLFile($zipOut, 'application/zip', 'images.zip');
  $postFields = [ 'file' => $cfile ];

  $up = curl_req('POST', $upUrl, $token, $postFields, true);
  $uploadOk = ($up['ok'] && $up['http'] === 200 && yahoo_xml_ok($up['body']));

  $submitSummary = [];
  $allSubmitOk = true;

  foreach ($itemCodes as $code) {
    $d = submit_one_item($sellerId, $code, $token);
    $submitSummary[] = $d;
    if (!$d['ok']) $allSubmitOk = false;
  }

  $okCount = 0; $ngCount = 0;
  foreach ($submitSummary as $d) {
    if (!empty($d['ok'])) $okCount++;
    else $ngCount++;
  }

  $elapsed = now_ms() - $t0;
  $anySubmitOk = ($okCount > 0);

  $finalOk = $strict
    ? ($deleteResult['ok'] && $uploadOk && $allSubmitOk)
    : ($uploadOk && $anySubmitOk);

  respond([
    'ok' => $finalOk,
    'action' => 'reupload_urls',
    'strict' => $strict,
    'run_id' => $runId,
    'saved_zip' => $zipOut,
    'item_codes' => $itemCodes,

    'notice' => [
      'count' => count($noticeSummary),
      'ok_count' => count(array_filter($noticeSummary, fn($x) => !empty($x['ok']))),
      'details_head' => array_slice($noticeSummary, 0, 20),
    ],

    'download' => [
      'count' => count($downloadSummary),
      'ok_count' => count(array_filter($downloadSummary, fn($x) => !empty($x['ok']))),
      'details_head' => array_slice($downloadSummary, 0, 80),
    ],

    'delete' => $deleteResult,

    'upload_pack' => [
      'http' => $up['http'],
      'ok'   => $uploadOk,
      'body_tail' => substr($up['body'], -800),
    ],

    'submit' => [
      'ok' => $allSubmitOk,
      'ok_count' => $okCount,
      'ng_count' => $ngCount,
      'details' => $submitSummary,
    ],

    'elapsed_ms' => $elapsed,
    'token_debug' => $debugTok,
    'v' => 'reupload_urls-v2',
  ], 200);
}

// =====================================================
// legacy: ZIP (upload/zip_url/zip_path) -> preprocess -> delete -> uploadPack -> [submit]
// v3: skip_submit=1 対応
// =====================================================

$zipMeta = [
  'source'   => null,
  'zip_url'  => null,
  'zip_path' => null,
  'upload'   => false,
];

if ($action === 'local_upload') {
  $itemCodes = parse_submit_item_codes($_POST);
  if (!$itemCodes) {
    respond(['ok' => false, 'error' => 'item_codes required'], 400);
  }

  $imgBase = '/var/www/html/img';
  
  safe_mkdir(RUN_BASE_DIR);
  $runId  = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
  $runDir = rtrim(RUN_BASE_DIR, '/') . '/' . $runId;
  safe_mkdir($runDir);
  $workDir = $runDir . '/work';
  safe_mkdir($workDir);

$foundCodes = [];
  foreach ($itemCodes as $code) {
    $code = trim($code);
    if ($code === '') continue;

    $srcDir = $imgBase . '/' . $code;
    if (!is_dir($srcDir)) continue;

    // .jpg と .jpeg の両方を探す
    $files = glob($srcDir . '/*.{jpg,jpeg,JPG,JPEG}', GLOB_BRACE);
    if (!$files) continue;

    foreach ($files as $src) {
      $fn = basename($src);
      // 拡張子を統一しつつ、大文字小文字はそのまま維持！
      $fn = preg_replace('/\.jpeg$/i', '.jpg', $fn);
      copy($src, $workDir . '/' . $fn);
    }
    $foundCodes[] = $code;
  }

  if (!$foundCodes) {
    respond(['ok' => false, 'error' => 'no images found on VPS', 'searched' => $imgBase], 404);
  }

  $cmd = PREPROCESS_BIN . ' ' . escapeshellarg($workDir) . ' 2>&1';
  $ppOut = []; $ppRet = 0;
  exec($cmd, $ppOut, $ppRet);

  $zipOut = $runDir . '/input_processed.zip';
  $zb = new ZipArchive();
  $zb->open($zipOut, ZipArchive::CREATE | ZipArchive::OVERWRITE);
  $added = 0;
  $dh = opendir($workDir);
  while (($fn = readdir($dh)) !== false) {
    if (!preg_match('/\.jpe?g$/i', $fn)) continue;
    $full = $workDir . '/' . $fn;
    if (is_file($full)) { $zb->addFile($full, $fn); $added++; }
  }
  closedir($dh);
  $zb->close();

  if ($added === 0) {
    respond(['ok' => false, 'error' => 'ZIP empty after preprocess'], 500);
  }

  $deleteResult = delete_images_for_codes($foundCodes, $sellerId, $token);

  $upUrl = EP_UPLOAD_PACK . '?seller_id=' . rawurlencode($sellerId);
  $cfile = new CURLFile($zipOut, 'application/zip', 'images.zip');
  $up = curl_req('POST', $upUrl, $token, ['file' => $cfile], true);
  $uploadOk = ($up['ok'] && $up['http'] === 200 && yahoo_xml_ok($up['body']));

  $skipSubmit = !empty($_REQUEST['skip_submit']);

  respond([
    'ok' => $uploadOk,
    'action' => 'local_upload',
    'run_id' => $runId,
    'item_codes' => $foundCodes,
    'added_jpg' => $added,
    'delete' => $deleteResult,
    'upload_pack' => [
      'http' => $up['http'],
      'ok' => $uploadOk,
      'body_tail' => substr($up['body'], -800),
    ],
    'skip_submit' => $skipSubmit,
  ], 200);
}

[$srcZip, $zipResolveErr] = resolve_zip_source($zipMeta);

if ($zipResolveErr) {
  respond([
    'ok' => false,
    'error' => $zipResolveErr,
    'hint' => 'Send zip as multipart (zip) OR send zip_url (https://img.niyantarose.com/dl/xxx.zip) OR send zip_path (allowed dirs only).',
    'files' => array_keys($_FILES ?? []),
    'post_keys' => array_keys($_POST ?? []),
  ], 400);
}

if (!$srcZip) {
  respond([
    'ok' => false,
    'error' => 'zip not provided for legacy action',
    'hint' => 'Legacy action requires a ZIP file.',
    'post_keys' => array_keys($_POST ?? []),
  ], 400);
}

clearstatcache(true, $srcZip);
$srcSize = @filesize($srcZip);

if ($srcSize === false || $srcSize <= 0) {
  respond(['ok' => false, 'error' => 'zip size=0', 'zipPath' => $srcZip, 'zip_meta' => $zipMeta], 400);
}

if ($srcSize > ZIP_LIMIT_BYTES) {
  respond([
    'ok' => false,
    'error' => 'zip too large',
    'size' => $srcSize,
    'limit' => ZIP_LIMIT_BYTES,
    'zip_meta' => $zipMeta,
  ], 400);
}

// run dir
safe_mkdir(RUN_BASE_DIR);
$runId  = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
$runDir = rtrim(RUN_BASE_DIR, '/') . '/' . $runId;
safe_mkdir($runDir);

$zipPath = $runDir . '/input.zip';

// --- 1) ZIP を用意 ---
if ($zipMeta['source'] === 'upload') {
  if (!isset($_FILES['zip'])) respond(['ok'=>false,'error'=>'upload zip missing unexpectedly'], 500);
  $f = $_FILES['zip'];
  if (($f['error'] ?? UPLOAD_ERR_OK) !== UPLOAD_ERR_OK) respond(['ok'=>false,'error'=>'UPLOAD_ERR','code'=>$f['error']], 400);
  if (!move_uploaded_file($f['tmp_name'], $zipPath)) respond(['ok'=>false,'error'=>'move_uploaded_file failed'], 500);
} else {
  if (!@copy($srcZip, $zipPath)) respond(['ok'=>false,'error'=>'copy zip failed','src'=>$srcZip], 500);
}

// --- 2) preprocess: unzip -> resize -> rezip ---
if (!is_file(PREPROCESS_BIN)) {
  respond(['ok'=>false,'error'=>'missing preprocess bin','bin'=>PREPROCESS_BIN], 500);
}

$workDir = $runDir . '/work';
safe_mkdir($workDir);

$za2 = new ZipArchive();
if ($za2->open($zipPath) !== true) {
  respond(['ok' => false, 'error' => 'ZIP_OPEN_FAIL_FOR_PREPROCESS'], 400);
}
$za2->extractTo($workDir);
$za2->close();

$cmd = PREPROCESS_BIN . ' ' . escapeshellarg($workDir) . ' 2>&1';
$ppOut = [];
$ppRet = 0;
exec($cmd, $ppOut, $ppRet);

// rezip jpg only
$zipOut = $runDir . '/input_processed.zip';
$zb = new ZipArchive();
if ($zb->open($zipOut, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
  respond(['ok' => false, 'error' => 'ZIP_CREATE_FAIL_FOR_PREPROCESS'], 500);
}

$added = 0;
$dh = opendir($workDir);
if ($dh !== false) {
  while (($fn = readdir($dh)) !== false) {
    if ($fn === '.' || $fn === '..') continue;
    if (!preg_match('/\.(jpe?g)$/i', $fn)) continue;
    $full = $workDir . '/' . $fn;
    if (is_file($full)) {
      $zb->addFile($full, $fn);
      $added++;
    }
  }
  closedir($dh);
}
$zb->close();

if ($ppRet !== 0 || $added === 0) {
  respond([
    'ok' => false,
    'error' => 'PREPROCESS_FAIL',
    'preprocess_ret' => $ppRet,
    'added_jpg' => $added,
    'preprocess_tail' => array_slice($ppOut, -80),
  ], 500);
}

$zipPath = $zipOut;

// --- 3) item_codes 決定 ---
$item_codes_source = 'zip_infer';

if (isset($_POST['item_codes_json']) && $_POST['item_codes_json'] !== '') {
  $item_codes_source = 'item_codes_json';

  $tmp = json_decode((string)$_POST['item_codes_json'], true);
  if (!is_array($tmp)) {
    respond([
      'ok' => false,
      'error' => 'invalid item_codes_json (must be JSON array)',
      'item_codes_source' => $item_codes_source,
      'raw_head' => substr((string)$_POST['item_codes_json'], 0, 200),
    ], 400);
  }

  $itemCodes = [];
  foreach ($tmp as $c) {
    $c = trim((string)$c);
    if ($c === '') continue;
    if (stripos($c, 'notice_') === 0) continue;
    $itemCodes[] = $c;
  }
  $itemCodes = array_values(array_unique($itemCodes));

  if (!$itemCodes) {
    respond([
      'ok' => false,
      'error' => 'item_codes_json became empty after normalize',
      'item_codes_source' => $item_codes_source,
    ], 400);
  }
} else {
  [$itemCodes, $zipErr] = parse_item_codes_from_zip($zipPath);
  if ($zipErr) respond(['ok' => false, 'error' => $zipErr], 400);
}

$t0 = now_ms();

// --- 4) DELETE (best-effort) ---
 $deleteResult = delete_images_for_codes($itemCodes, $sellerId, $token); 


// --- 5) UPLOAD PACK ---
$upUrl = EP_UPLOAD_PACK . '?seller_id=' . rawurlencode($sellerId);
$cfile = new CURLFile($zipPath, 'application/zip', 'images.zip');
$postFields = [ 'file' => $cfile ];

$up = curl_req('POST', $upUrl, $token, $postFields, true);
$uploadOk = ($up['ok'] && $up['http'] === 200 && yahoo_xml_ok($up['body']));

// --- 6) SUBMIT ---
// ▼ 修正：$_REQUESTに変更してGASのURLパラメータを受け取れるようにしました
$skipSubmit = !empty($_REQUEST['skip_submit']);
$submitSummary = [];
$allSubmitOk = true;
$okCount = 0;
$ngCount = 0;

if (!$skipSubmit) {
  foreach ($itemCodes as $code) {
    $d = submit_one_item($sellerId, $code, $token);
    $submitSummary[] = $d;
    if (!$d['ok']) $allSubmitOk = false;
  }
  foreach ($submitSummary as $d) {
    if (!empty($d['ok'])) $okCount++;
    else $ngCount++;
  }
}

$elapsed = now_ms() - $t0;

// ▼ 修正：こちらも$_REQUESTに変更
$strict = !empty($_REQUEST['strict']);

if ($skipSubmit) {
  $finalOk = $uploadOk;
} else {
  $anySubmitOk = ($okCount > 0);
  $finalOk = $strict
    ? ($deleteResult['ok'] && $uploadOk && $allSubmitOk)
    : ($uploadOk && $anySubmitOk);
}

// 修正：一番下にあった遅すぎるエラーチェックを削除済み

respond([
  'ok' => $finalOk,
  'action' => 'legacy',
  'strict' => $strict,
  'skip_submit' => $skipSubmit,
  'run_id' => $runId,
  'saved_zip' => $zipOut,
  'item_codes' => $itemCodes,
  'delete' => $deleteResult,
  'upload_pack' => [
    'http' => $up['http'],
    'ok'   => $uploadOk,
    'body_tail' => substr($up['body'], -800),
  ],
  'submit' => [
    'ok' => $skipSubmit ? null : $allSubmitOk,
    'skipped' => $skipSubmit,
    'ok_count' => $okCount,
    'ng_count' => $ngCount,
    'details' => $submitSummary,
  ],
  'elapsed_ms' => $elapsed,
  'token_debug' => $debugTok,
  'v' => 'legacy-fixed-v1',
], 200);
}
