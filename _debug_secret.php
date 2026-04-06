<?php
header('Content-Type: application/json; charset=utf-8');

$out = [
  'time' => date('c'),
  'remote_addr' => $_SERVER['REMOTE_ADDR'] ?? null,
  'has_HTTP_X_VPS_SECRET' => isset($_SERVER['HTTP_X_VPS_SECRET']),
  'HTTP_X_VPS_SECRET' => $_SERVER['HTTP_X_VPS_SECRET'] ?? null,
  'has_POST_secret' => isset($_POST['secret']),
  'POST_secret' => $_POST['secret'] ?? null,
  'has_GET_secret' => isset($_GET['secret']),
  'GET_secret' => $_GET['secret'] ?? null,
  'http_authorization' => $_SERVER['HTTP_AUTHORIZATION'] ?? null,
];
echo json_encode($out, JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES);
