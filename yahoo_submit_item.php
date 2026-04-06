<?php
header('Content-Type: application/json');

$token    = getenv('YAHOO_ACCESS_TOKEN');
$sellerId = getenv('YAHOO_SELLER_ID');

$url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem';

$post = http_build_query([
  'seller_id' => $sellerId,
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
