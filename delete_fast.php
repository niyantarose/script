<?php
// Yahoo一括削除爆速版 (delete_fast.php)

// 「外部から読み込まれている」という合図を定義
define('IS_INCLUDED', true);

// 既存の機能を読み込む
require_once __DIR__ . '/yahoo_zip_upload.php';


header('Content-Type: application/json');

// GASからの合言葉チェック
require_secret_or_403();

$seller_id = $_POST['seller_id']    ?? '';
$item_codes_raw = $_POST['item_codes'] ?? '';

if (!$seller_id || !$item_codes_raw) {
    echo json_encode(['status' => 'error', 'message' => 'Missing parameters']);
    exit;
}

// トークンを yahoo_zip_upload.php の関数で取得
$debugTok = [];
$access_token = refresh_access_token_from_env($debugTok);

if (!$access_token) {
    echo json_encode(['status' => 'error', 'message' => 'Token refresh failed', 'debug' => $debugTok]);
    exit;
}

$item_codes = explode(',', $item_codes_raw);
$delete_ids = [];

foreach ($item_codes as $code) {
    $code = trim($code);
    if (!$code) continue;
    // メイン + 1~20枚を削除リストに追加
    $delete_ids[] = $code;
    for ($i = 1; $i <= 20; $i++) { $delete_ids[] = "{$code}_{$i}"; }
}

$delete_ids = array_unique($delete_ids);

// 100件ずつ一気に削除 (ここが速さの秘密)
$results = [];
foreach (array_chunk($delete_ids, 100) as $chunk) {
    $ch = curl_init("https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage");
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, ["Authorization: Bearer $access_token"]);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query(['seller_id' => $seller_id, 'image_id' => implode(',', $chunk)]));
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    $results[] = curl_exec($ch);
    curl_close($ch);
}

echo json_encode(['status' => 'success', 'deleted_count' => count($delete_ids)]);
