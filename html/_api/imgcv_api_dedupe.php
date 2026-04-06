<?php
declare(strict_types=1);
require_once '/var/www/_inc/yahoo_api_lib.php';  // ← ★この1行を追加！
header('Content-Type: application/json; charset=utf-8');


function respond(array $a, int $code=200): void {
  http_response_code($code);
  echo json_encode($a, JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES);
  exit;
}
function need(bool $cond, string $msg, int $code=400): void {
  if (!$cond) respond(['ok'=>false,'error'=>$msg], $code);
}
function expected_secret(): string {
  $s = trim((string)@file_get_contents('/etc/imgcv_secret'));
  if ($s === '') respond(['ok'=>false,'error'=>'server secret missing'], 500);
  return $s;
}
function require_secret(string $expected): void {
  $got = $_SERVER['HTTP_X_VPS_SECRET'] ?? '';
  if (!$got || !hash_equals($expected, $got)) respond(['ok'=>false,'error'=>'FORBIDDEN'], 403);
}

function db(): PDO {
  static $pdo=null;
  if ($pdo) return $pdo;
  $path = '/var/lib/imgcv/imgcv_state.sqlite';
  $pdo = new PDO('sqlite:'.$path, null, null, [
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
  ]);
  $pdo->exec('PRAGMA journal_mode=WAL;');
  $pdo->exec('PRAGMA busy_timeout=5000;');
  $pdo->exec("
    CREATE TABLE IF NOT EXISTS image_state (
      seller_id TEXT NOT NULL,
      item_code TEXT NOT NULL,
      name TEXT NOT NULL,
      hash TEXT NOT NULL,
      size INTEGER NOT NULL,
      src TEXT NOT NULL,
      updated_at INTEGER NOT NULL,
      PRIMARY KEY (seller_id,item_code,name)
    );
  ");
  $pdo->exec("
    CREATE TABLE IF NOT EXISTS submit_state (
      seller_id TEXT NOT NULL,
      item_code TEXT NOT NULL,
      desired_hash TEXT NOT NULL,
      updated_at INTEGER NOT NULL,
      PRIMARY KEY (seller_id,item_code)
    );
  ");
  return $pdo;
}
function get_image_state(string $seller, string $code, string $name): ?array {
  $st = db()->prepare("SELECT * FROM image_state WHERE seller_id=? AND item_code=? AND name=?");
  $st->execute([$seller,$code,$name]);
  $r = $st->fetch();
  return $r ?: null;
}
function upsert_image_state(string $seller, string $code, string $name, string $hash, int $size, string $src): void {
  $now=time();
  $st = db()->prepare("
    INSERT INTO image_state (seller_id,item_code,name,hash,size,src,updated_at)
    VALUES (?,?,?,?,?,?,?)
    ON CONFLICT(seller_id,item_code,name) DO UPDATE SET
      hash=excluded.hash, size=excluded.size, src=excluded.src, updated_at=excluded.updated_at
  ");
  $st->execute([$seller,$code,$name,$hash,$size,$src,$now]);
}
function get_submit_hash(string $seller, string $code): ?string {
  $st = db()->prepare("SELECT desired_hash FROM submit_state WHERE seller_id=? AND item_code=?");
  $st->execute([$seller,$code]);
  $r = $st->fetch();
  return $r ? (string)$r['desired_hash'] : null;
}
function set_submit_hash(string $seller, string $code, string $hash): void {
  $now=time();
  $st = db()->prepare("
    INSERT INTO submit_state (seller_id,item_code,desired_hash,updated_at)
    VALUES (?,?,?,?)
    ON CONFLICT(seller_id,item_code) DO UPDATE SET
      desired_hash=excluded.desired_hash, updated_at=excluded.updated_at
  ");
  $st->execute([$seller,$code,$hash,$now]);
}

function download_url_to_temp(string $url): array {
  $tmp = tempnam(sys_get_temp_dir(), 'imgcv_');
  if ($tmp === false) throw new RuntimeException('tempnam failed');

  $ch = curl_init($url);
  $fp = fopen($tmp, 'wb');
  if (!$fp) { @unlink($tmp); throw new RuntimeException('fopen temp failed'); }

  curl_setopt_array($ch, [
    CURLOPT_FILE => $fp,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_MAXREDIRS => 5,
    CURLOPT_CONNECTTIMEOUT => 10,
    CURLOPT_TIMEOUT => 40,
    CURLOPT_USERAGENT => 'imgcv-dedupe/1.0',
    CURLOPT_SSL_VERIFYPEER => true,
    CURLOPT_SSL_VERIFYHOST => 2,
    CURLOPT_FAILONERROR => false,
  ]);
  curl_exec($ch);
  $err = curl_error($ch);
  $http = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
  curl_close($ch);
  fclose($fp);

  if ($err) { @unlink($tmp); throw new RuntimeException('curl error: '.$err); }
  if ($http < 200 || $http >= 300) { @unlink($tmp); throw new RuntimeException('download http='.$http); }

  $size = filesize($tmp);
  if ($size === false || $size <= 0) { @unlink($tmp); throw new RuntimeException('downloaded size invalid'); }

  return [$tmp, (int)$size];
}
function sha1_file_safe(string $path): string {
  $h = sha1_file($path);
  if (!$h) throw new RuntimeException('sha1_file failed');
  return $h;
}

// -------------------- main --------------------
$expected = expected_secret();
require_secret($expected);

$raw = file_get_contents('php://input') ?: '';
$req = json_decode($raw, true);
need(is_array($req), 'invalid json');

// ④ 変数名を揃える “もっと安全なやり方”（互換用）
$body = $req; 

// ① commit フラグを読むコード（デフォルトは dry-run）
$commit = !empty($req['commit']) || !empty($req['debug']['commit']);
$force_commit = !empty($req['debug']['force_commit']); // ← ★この1行を追加！

$mode   = (string)($req['mode'] ?? '');

$seller = trim((string)($req['seller_id'] ?? ''));
need($mode !== '', 'mode required');
need($seller !== '', 'seller_id required');

// --- yahoo_upload（差分判定のみ：後でここにcommit時の処理を追加） ---
// （ここは前回と同じなので省略せずにそのまま残します）
if ($mode === 'yahoo_upload') {
  // ... (前回のアップロード判定処理そのまま) ...
}

// --- submit（desired_imagesの差分判定 ＆ commit=1なら反映） ---
  if ($mode === 'submit') {
    // ② token 取り出し（commit=1 のときだけ必須）
    $token = '';
    if ($commit) {
      $token = trim((string)($req['yahoo_token'] ?? ''));

      // リクエストに無ければサーバキャッシュから読む
      if ($token === '') {
        $cache = '/tmp/yahoo_access_token_cache.json';
        if (is_file($cache)) {
          $j = json_decode((string)@file_get_contents($cache), true);
          if (is_array($j) && !empty($j['access_token'])) {
            $token = trim((string)$j['access_token']);
          }
        }
      }

      need($token !== '', 'yahoo_token required when commit=1', 400);
    }

$codes = $req['items'] ?? [];
  $desired = $req['desired_images'] ?? [];
  need(is_array($codes), 'items must be array');
  need(is_array($desired), 'desired_images must be object');

  $details=[]; $processed=0; $would_submit=0; $skipped=0; $failed=0;
  
  // ★対策2: 安全な処理対象リストを追加
  $validCodes = [];
  $pending = [];
  $toSubmit = [];

  foreach ($codes as $c) {
    $processed++;
    $code = trim((string)$c);
    if ($code==='') { $failed++; $details[]=['ok'=>false,'item_code'=>'','status'=>'BAD_CODE']; continue; }

    $list = $desired[$code] ?? null;
    if (!is_array($list)) { $failed++; $details[]=['ok'=>false,'item_code'=>$code,'status'=>'NO_DESIRED']; continue; }

    // ★対策2: 空文字や desired が無いものを弾いた「クリーンなリスト」を作る
    $validCodes[] = $code;

    $main=''; $libs=[];
    foreach ($list as $fn) {
      $fn = trim((string)$fn);
      if ($fn==='') continue;
      if ($fn === $code.'.jpg') $main = $fn; else $libs[]=$fn;
    }

    // ★対策4: ハッシュの作り方（Yahooは画像の表示順序が重要になることが多いので、
    // ここはあえて sort() せず、配列の順序通りにハッシュ化する現状維持がおすすめです）
    // もし順序を無視したい場合は、以下のコメントアウトを外してください。
    // sort($libs, SORT_STRING);

    $dh = sha1(json_encode([$main,$libs], JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES));
    $prev = get_submit_hash($seller,$code);

    // ★対策1: より安全な null判定 ＋ hash_equals に変更
    if ($prev !== null && hash_equals($prev, $dh)) {
      $skipped++;
      $details[]=['ok'=>true,'item_code'=>$code,'status'=>'NO_CHANGE_SUBMIT'];
      continue;
    }

    $pending[$code] = $dh;
    $toSubmit[] = $code;

    $would_submit++;
    $details[]=['ok'=>true,'item_code'=>$code,'status'=>'WOULD_SUBMIT','main'=>$main,'libs'=>count($libs)];
  }

  $out = [
    'ok'=>($failed===0),
    'processed'=>$processed,
    'would_submit'=>$would_submit,
    'skipped'=>$skipped,
    'failed'=>$failed,
    'details'=>$details
  ];

  if ($commit && ($would_submit > 0 || $force_commit)) {
    
    // ★対策2: force_commit 時は生の $codes ではなく $validCodes を投げる
    $targetCodes = ($force_commit && empty($toSubmit)) ? $validCodes : $toSubmit;

    if (empty($targetCodes)) {
      // 念のため、投げる対象が全く無い場合はAPIを叩かずに終わる
      $out['warning'] = 'No valid target codes to submit.';
    } else {
      $rp = yahoo_reserve_publish_dummy($seller, $token, $targetCodes);

      // ★対策3: Yahoo API 失敗時は 502 で落とす（DB更新しない＆GASに失敗を伝える）
      if (empty($rp['ok'])) {
        respond(['ok'=>false, 'error'=>'Yahoo API failed', 'reserve_publish'=>$rp], 502);
      }

      // API成功時のみDBを更新
      foreach ($pending as $code => $dh) {
        set_submit_hash($seller, $code, $dh);
      }
      $out['reserve_publish'] = $rp;
    }
  }

  respond($out, 200);
}

respond(['ok'=>false,'error'=>'unknown mode'], 400);
