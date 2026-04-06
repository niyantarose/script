<?php
header('Content-Type: application/json');

// ==================================================
// Bearer（GAS側で管理）
// ==================================================
$auth = $_SERVER['HTTP_AUTHORIZATION'] ?? '';
if (!preg_match('/^Bearer\s+(.+)$/', $auth, $m)) {
  http_response_code(401);
  echo json_encode(['ok'=>false,'error'=>'missing bearer']);
  exit;
}
$accessToken = $m[1];

// 安全装置（短すぎるトークンを弾く）
if (strlen($accessToken) < 50) {
  http_response_code(401);
  echo json_encode(['ok'=>false,'error'=>'invalid bearer']);
  exit;
}

// ==================================================
// 入力
// ==================================================
$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

$itemCode = $data['item_code'] ?? null;
$zipPath  = $data['zip_path']  ?? null;

if (!$itemCode || !$zipPath || !is_file($zipPath)) {
  http_response_code(400);
  echo json_encode([
    'ok'=>false,
    'error'=>'missing item_code or zip_path',
    'parsed'=>$data
  ]);
  exit;
}

// ==================================================
// 設定
// ==================================================
$sellerId = 'niyantarose';

// ==================================================
// 共通POST
// ==================================================
function curlPost($url, $post, $token) {
  $ch = curl_init($url);
  curl_setopt_array($ch, [
    CURLOPT_POST           => true,
    CURLOPT_POSTFIELDS     => $post,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HTTPHEADER     => [
      'Authorization: Bearer ' . $token,
    ],
    CURLOPT_TIMEOUT        => 120,
  ]);
  $res = curl_exec($ch);
  $err = curl_error($ch);
  curl_close($ch);
  if ($err) throw new Exception($err);
  return $res;
}

// ==================================================
// 1. 全削除
// ==================================================
$resDelete = curlPost(
  'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage',
  http_build_query([
    'seller_id' => $sellerId,
    'item_code' => $itemCode, // ★ 全削除
  ]),
  $accessToken
);

// ==================================================
// 2. ZIPアップロード
// ==================================================
$uploadUrl = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImagePack?seller_id={$sellerId}";
$cfile = new CURLFile($zipPath, 'application/zip', basename($zipPath));

$resUpload = curlPost(
  $uploadUrl,
  ['file' => $cfile],
  $accessToken
);

// ==================================================
// 3. submit
// ==================================================
$resSubmit = curlPost(
  'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem',
  http_build_query([
    'seller_id' => $sellerId,
    'item_code' => $itemCode,
  ]),
  $accessToken
);

// ==================================================
// 完了
// ==================================================
error_log("[YAHOO_EXEC] {$itemCode} done");

echo json_encode([
  'ok'     => true,
  'item'   => $itemCode,
  'delete' => $resDelete,
  'upload' => $resUpload,
  'submit' => $resSubmit,
], JSON_UNESCAPED_UNICODE);
exit;
