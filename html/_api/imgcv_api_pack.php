<?php
// /var/www/html/_api/imgcv_api_pack.php
// GASから codes + yahoo_token を受け取り、VPS画像をZIP→uploadItemImagePack→submitItem
header('Content-Type: application/json; charset=UTF-8');

require_once '/var/www/_inc/yahoo_api_lib.php';

function respond(array $a, int $code = 200): void {
    http_response_code($code);
    echo json_encode($a, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

// ── secret 認証 ──────────────────────────────
$serverSecret = trim((string)@file_get_contents('/etc/niyantarose/ssh_secret.txt'));
$clientSecret = trim((string)($_SERVER['HTTP_X_VPS_SECRET'] ?? ''));
if ($serverSecret === '' || $clientSecret === '' || !hash_equals($serverSecret, $clientSecret)) {
    respond(['ok' => false, 'error' => 'FORBIDDEN'], 403);
}

// ── リクエスト ────────────────────────────────
$req = json_decode((string)file_get_contents('php://input'), true);
if (!is_array($req)) respond(['ok' => false, 'error' => 'BAD_JSON'], 400);

$seller = trim((string)($req['seller_id']    ?? ''));
$token  = trim((string)($req['yahoo_token']  ?? ''));
$codes  = $req['items'] ?? [];
$commit = !empty($req['commit']);

if ($seller === '') respond(['ok' => false, 'error' => 'seller_id required'], 400);
if (!is_array($codes) || count($codes) === 0) respond(['ok' => false, 'error' => 'items required'], 400);

// ── トークンをキャッシュに保存 ─────────────────
if ($token !== '') {
    $cache = '/tmp/yahoo_access_token_cache.json';
    @file_put_contents($cache, json_encode([
        'access_token' => $token,
        'saved_at'     => time(),
        'expires_at'   => time() + 3600,
    ]));
} else {
    // キャッシュから読む
    $cache = '/tmp/yahoo_access_token_cache.json';
    if (is_file($cache)) {
        $j = json_decode((string)@file_get_contents($cache), true);
        if (is_array($j) && !empty($j['access_token'])) {
            $token = trim((string)$j['access_token']);
        }
    }
}

if ($commit && $token === '') respond(['ok' => false, 'error' => 'yahoo_token required'], 400);

// ── トークン事前検証 ──────────────────────────
if ($commit) {
    $tokenCheck = yahoo_validate_token($seller, $token);
    if (!$tokenCheck['ok']) {
        respond(['ok' => false, 'error' => 'token_expired', 'detail' => $tokenCheck['detail']], 401);
    }
}

// ── 定数 ─────────────────────────────────────
define('IMG_ROOT',    '/var/www/html/img');
define('MAX_ZIP_MB',  24 * 1024 * 1024);  // 24MB（25MB制限の余裕）
define('MAX_ZIP_FILES', 35);
define('UPLOAD_URL', 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImagePack');
define('SUBMIT_URL', 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem');

// ── 画像収集 ─────────────────────────────────
$allFiles = [];  // [['path'=>..., 'name'=>...], ...]
$missing  = [];

foreach ($codes as $code) {
    $code = trim((string)$code);
    if ($code === '') continue;

    $dir = IMG_ROOT . '/' . $code;
    if (!is_dir($dir)) { $missing[] = $code; continue; }

    // main: code.jpg
    $main = $dir . '/' . $code . '.jpg';
    if (is_file($main)) {
        $allFiles[] = ['path' => $main, 'name' => $code . '.jpg', 'code' => $code];
    }

    // libs: code_1.jpg 〜 code_20.jpg
    for ($i = 1; $i <= 20; $i++) {
        $lib = $dir . '/' . $code . '_' . $i . '.jpg';
        if (is_file($lib)) {
            $allFiles[] = ['path' => $lib, 'name' => $code . '_' . $i . '.jpg', 'code' => $code];
        }
    }
}

if (count($allFiles) === 0) {
    respond([
        'ok'      => false,
        'error'   => 'no_images_found',
        'missing' => $missing,
        'codes'   => $codes,
    ]);
}

// ── ZIP分割・アップロード ──────────────────────
$results    = [];  // code => ['ok'=>bool, 'status'=>string]
$uploadLogs = [];

if ($commit) {
    $chunks = [];
    $chunk  = [];
    $chunkSize = 0;

    foreach ($allFiles as $f) {
        $size = filesize($f['path']);
        if (
            count($chunk) > 0 && (
                $chunkSize + $size > MAX_ZIP_MB ||
                count($chunk) >= MAX_ZIP_FILES
            )
        ) {
            $chunks[] = $chunk;
            $chunk = [];
            $chunkSize = 0;
        }
        $chunk[] = $f;
        $chunkSize += $size;
    }
    if (count($chunk) > 0) $chunks[] = $chunk;

    foreach ($chunks as $ci => $files) {
        // ZIP作成
        $zipPath = '/tmp/yahoo_pack_' . getmypid() . '_' . $ci . '.zip';
        $zip = new ZipArchive();
        if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            $uploadLogs[] = ['chunk' => $ci, 'ok' => false, 'error' => 'zip_create_failed'];
            continue;
        }
        foreach ($files as $f) {
            $zip->addFile($f['path'], $f['name']);
        }
        $zip->close();

        // uploadItemImagePack
        $url = UPLOAD_URL . '?seller_id=' . urlencode($seller);
        $ch = curl_init($url);
        curl_setopt_array($ch, [
            CURLOPT_POST           => true,
            CURLOPT_POSTFIELDS     => ['file' => new CURLFile($zipPath, 'application/zip', 'images.zip')],
            CURLOPT_HTTPHEADER     => ['Authorization: Bearer ' . $token],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 120,
        ]);
        $body = (string)curl_exec($ch);
        $http = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
        curl_close($ch);
        @unlink($zipPath);

        $isOk = ($http >= 200 && $http < 300 && strpos($body, '<Error>') === false);
        $uploadLogs[] = [
            'chunk'  => $ci + 1,
            'files'  => count($files),
            'http'   => $http,
            'ok'     => $isOk,
            'body'   => substr($body, 0, 200),
        ];

        error_log('DEBUG pack chunk=' . ($ci+1) . ' http=' . $http . ' files=' . count($files));

        if ($isOk) {
            foreach ($files as $f) {
                $results[$f['code']] = ['ok' => true, 'status' => 'UPLOADED'];
            }
        } else {
            foreach ($files as $f) {
                $results[$f['code']] = ['ok' => false, 'status' => 'UPLOAD_FAILED'];
            }
        }
    }

    // submitItem（アップ成功コードのみ）
    $submitResults = [];
foreach ($results as $code => $r) {
        if (!$r['ok']) continue;

        $shttp = 0;
        $sbody = '';
        $submitOk = false;

        for ($retry = 0; $retry < 5; $retry++) {
            $ch = curl_init(SUBMIT_URL);
            curl_setopt_array($ch, [
                CURLOPT_POST           => true,
                CURLOPT_POSTFIELDS     => http_build_query([
                    'seller_id' => $seller,
                    'item_code' => $code,
                ]),
                CURLOPT_HTTPHEADER     => [
                    'Authorization: Bearer ' . $token,
                    'Content-Type: application/x-www-form-urlencoded',
                ],
                CURLOPT_RETURNTRANSFER => true,
                CURLOPT_TIMEOUT        => 30,
            ]);
            $sbody = (string)curl_exec($ch);
            $shttp = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
            curl_close($ch);

            $submitOk = ($shttp >= 200 && $shttp < 300 && strpos($sbody, '<Error>') === false);
            if ($submitOk || strpos($sbody, 'ed-00006') === false) break;
            error_log('DEBUG submitItem ed-00006 retry=' . $retry . ' code=' . $code);
            sleep(2);
        }

        $submitResults[$code] = ['http' => $shttp, 'ok' => $submitOk, 'body' => substr($sbody, 0, 300)];
        error_log('DEBUG submitItem code=' . $code . ' http=' . $shttp);
    }


}

respond([
    'ok'             => true,
    'commit'         => $commit,
    'total_files'    => count($allFiles),
    'total_codes'    => count($codes),
    'missing'        => $missing,
    'upload_logs'    => $uploadLogs,
    'results'        => $results,
    'submit_results' => $submitResults ?? [],
]);
