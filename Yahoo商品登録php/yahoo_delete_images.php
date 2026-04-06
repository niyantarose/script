<?php
header('Content-Type: application/json');

$token    = getenv('YAHOO_ACCESS_TOKEN');
$sellerId = getenv('YAHOO_SELLER_ID');

$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

$itemCode = $data['item_code'] ?? null;
if (!$itemCode) {
  http_response_code(400);
  echo json_encode(['ok'=>false,'error'=>'missing item_code']);
  exit;
}

$url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage';

$post = http_build_query([
  'seller_id' => $sellerId,
  'image_id'  => $itemCode, // ★ 全削除
]);

$ch = curl_init($url);
curl_setopt_array($ch, [
  CURLOPT_POST           => true,
  CURLOPT_POSTFIELDS     => $post,
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_HTTPHEADER     => [
    "Authorization: Bearer $token",
  ],
]);

$res = curl_exec($ch);
curl_close($ch);

echo json_encode([
  'ok' => true,
  'response' => $res,
]);

