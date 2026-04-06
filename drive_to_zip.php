<?php
// drive_to_zip.php（公開ZIP生成 + キャッシュ + 安全対策 + 自動掃除）
header('Content-Type: application/json; charset=utf-8');
require_once __DIR__ . '/vendor/autoload.php';

// ▼環境設定
const KEY_FILE_PATH = '/var/www/secret/service-account.json';
const WORK_ROOT     = '/var/www/private/temp_work';
const CACHE_DIR     = '/var/www/private/drive_cache';
const TOKEN_DIR     = '/var/www/private/tokens';
const PUBLISH_DIR   = '/var/www/html/dl';

// ---------- util ----------
function safe_name($s): string {
    $s = (string)$s;
    $s = preg_replace('/[^A-Za-z0-9._-]+/', '_', $s);
    $s = trim($s, '._-');
    return $s === '' ? 'NONAME' : $s;
}
function ensure_dir($path, $mode): void {
    if (file_exists($path) && !is_dir($path)) {
        throw new Exception("Not a directory: {$path}");
    }
    if (!file_exists($path)) {
        if (!mkdir($path, $mode, true)) {
            throw new Exception("Directory creation failed: {$path}");
        }
    }
}
function cleanup_old_zips($dir, $days = 2): void {
    $limit = time() - ($days * 86400);
    foreach (glob($dir . '/*.zip') ?: [] as $f) {
        if (is_file($f) && filemtime($f) < $limit) @unlink($f);
    }
}
function cleanup_once_in_a_while($publishDir, $tokenDir, $days = 2, $intervalSec = 3600): void {
    $stamp = rtrim($tokenDir, '/') . '/last_cleanup.txt';
    $now = time();
    $last = is_file($stamp) ? (int)@file_get_contents($stamp) : 0;
    if ($now - $last < $intervalSec) return;

    // なるべく競合しないようにLOCK_EXで更新
    @file_put_contents($stamp, (string)$now, LOCK_EX);
    cleanup_old_zips($publishDir, $days);
}
function atomic_publish_zip($srcZip, $dstZip): void {
    // 同一FSならrenameが最速＆原子的。失敗時はcopy→unlinkでフォールバック。
    $tmp = $dstZip . '.tmp';
    if (file_exists($tmp)) @unlink($tmp);
    if (file_exists($dstZip)) @unlink($dstZip);

    if (@rename($srcZip, $tmp)) {
        if (!@rename($tmp, $dstZip)) {
            // 最後のrenameが失敗したら戻すのは難しいので、tmp残骸を消す
            @unlink($tmp);
            throw new Exception("Publish rename failed: {$tmp} -> {$dstZip}");
        }
        return;
    }
    // rename失敗 → copy fallback
    if (!@copy($srcZip, $tmp)) throw new Exception("Publish copy failed: {$srcZip} -> {$tmp}");
    @unlink($srcZip);
    if (!@rename($tmp, $dstZip)) {
        @unlink($tmp);
        throw new Exception("Publish finalize failed: {$tmp} -> {$dstZip}");
    }
}

// ---------- input ----------
$raw = file_get_contents('php://input');
$payload = json_decode($raw, true);
if (!is_array($payload) || !isset($payload['items']) || !is_array($payload['items'])) {
    http_response_code(400);
    echo json_encode(['ok' => false, 'error' => 'Invalid JSON'], JSON_UNESCAPED_SLASHES);
    exit;
}

// ---------- main ----------
try {
    ensure_dir(WORK_ROOT, 0770);
    ensure_dir(CACHE_DIR, 0770);
    ensure_dir(TOKEN_DIR, 0770);
    ensure_dir(PUBLISH_DIR, 0755);

    // Google Drive client
    $client = new Google\Client();
    $client->setAuthConfig(KEY_FILE_PATH);
    $client->addScope(Google\Service\Drive::DRIVE_READONLY);
    $drive = new Google\Service\Drive($client);

    // 実行ごとにユニークな作業ディレクトリ（並列でも衝突しにくい）
    $runId  = date('Ymd_His') . '_' . bin2hex(random_bytes(4));
    $runDir = rtrim(WORK_ROOT, '/') . '/run_' . $runId;
    ensure_dir($runDir, 0770);

    $results = [];
    $stats = ['cached_hit' => 0, 'downloaded' => 0, 'converted' => 0, 'files_in_zip' => 0];

    // publicZipName（テスト用）: itemsが1個のときだけ上書き可能
    $publicZipName = '';
    if (count($payload['items']) === 1 && isset($payload['publicZipName'])) {
        $publicZipName = safe_name($payload['publicZipName']);
    }

    foreach ($payload['items'] as $item) {
        $code  = safe_name($item['code'] ?? '');
        $files = is_array($item['files'] ?? null) ? $item['files'] : [];

        // ZIP名（通常は code.zip / テスト時だけ publicZipName.zip）
        $zipBase = ($publicZipName !== '') ? $publicZipName : $code;
        $zipName = $zipBase . '.zip';

        $productDir = $runDir . '/' . $code;
        ensure_dir($productDir, 0770);

        try {
            // --- download & convert ---
            foreach ($files as $f) {
                $fileId = (string)($f['fileId'] ?? '');
                if ($fileId === '') throw new Exception("Missing fileId (code={$code})");

                $name = safe_name($f['zipName'] ?? 'noname.jpg');
                $name = preg_replace('/\.(webp|png|jpeg|gif)$/i', '.jpg', $name);
                if (!str_ends_with(strtolower($name), '.jpg')) $name .= '.jpg';

                $cachePath = rtrim(CACHE_DIR, '/') . '/' . $fileId . '.jpg';
                $savePath  = $productDir . '/' . $name;

                if (is_file($cachePath)) {
                    if (!copy($cachePath, $savePath)) throw new Exception("Cache copy failed: {$cachePath}");
                    $stats['cached_hit']++;
                    continue;
                }

                // Drive download
                $resp = $drive->files->get($fileId, ['alt' => 'media', 'supportsAllDrives' => true]);
                $bin  = $resp->getBody()->getContents();
                $stats['downloaded']++;

                // decode & save JPEG
                $img = @imagecreatefromstring($bin);

                // GDで無理なら Imagick があれば救う（あれば強い）
                if ($img === false && class_exists('Imagick')) {
                    $im = new Imagick();
                    $im->readImageBlob($bin);
                    $im->setImageFormat('jpeg');
                    $im->setImageCompressionQuality(90);
                    $im->writeImage($cachePath);
                    $im->clear();
                    $im->destroy();
                    $stats['converted']++;

                    if (!copy($cachePath, $savePath)) throw new Exception("File copy failed: {$cachePath}");
                    continue;
                }

                if ($img === false) {
                    throw new Exception("Image decode failed (GD/Imagick) fileId={$fileId}");
                }

                if (!imagejpeg($img, $cachePath, 90)) {
                    imagedestroy($img);
                    throw new Exception("Image save failed: {$cachePath}");
                }
                imagedestroy($img);
                $stats['converted']++;

                if (!copy($cachePath, $savePath)) throw new Exception("File copy failed: {$cachePath}");
            }

            // --- zip ---
            $zipTmp = $runDir . '/' . $zipName; // まず非公開側で作る
            $zip = new ZipArchive();
            if ($zip->open($zipTmp, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
                throw new Exception("ZIP creation failed: {$zipTmp}");
            }

            $added = 0;
            foreach (glob($productDir . '/*') ?: [] as $p) {
                if (is_file($p)) {
                    $zip->addFile($p, basename($p));
                    $added++;
                }
            }
            $zip->close();

            if ($added === 0) {
                @unlink($zipTmp);
                throw new Exception("ZIP has no files (code={$code})");
            }
            $stats['files_in_zip'] += $added;

            // --- publish to /dl ---
            $publicPath = rtrim(PUBLISH_DIR, '/') . '/' . $zipName;
            atomic_publish_zip($zipTmp, $publicPath);

            $results[] = [
                'code'   => $code,
                'zip'    => $zipName,
                'zipUrl' => 'https://img.niyantarose.com/dl/' . rawurlencode($zipName),
                'status' => 'success',
            ];
        } finally {
            // productDir cleanup
            foreach (glob($productDir . '/*') ?: [] as $x) {
                if (is_file($x)) @unlink($x);
            }
            @rmdir($productDir);
        }
    }

    // runDir cleanup
    @rmdir($runDir);

    // ZIP掃除：毎回やると重いので、だいたい1時間に1回だけ
    cleanup_once_in_a_while(PUBLISH_DIR, TOKEN_DIR, 2, 3600);

    echo json_encode([
        'ok'      => true,
        'results' => $results,
        'stats'   => $stats,
        'version' => '20260209_01'
    ], JSON_UNESCAPED_SLASHES);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['ok' => false, 'error' => $e->getMessage()], JSON_UNESCAPED_SLASHES);
}
