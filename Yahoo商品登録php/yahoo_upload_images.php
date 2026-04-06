<?php
header('Content-Type: application/json');

$token    = getenv('YAHOO_ACCESS_TOKEN');
$sellerId = getenv('YAHOO_SELLER_ID');

$zipPath = $_POST['zip_path'] ?? null;
if (!$zipPath || !is_file($zipPath)) {
  http_response_code(400);
  echo json_encode(['ok'=>false,'error'=>'zip not found']);
  exit;
}

$url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImagePack?seller_id=$sellerId";

$cfile = new CURLFile($zipPath, 'application/zip', basename($zipPath));

$ch = curl_init($url);
curl_setopt_array($ch, [
  CURLOPT_POST           => true,
  CURLOPT_POSTFIELDS     => ['file' => $cfile],
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
