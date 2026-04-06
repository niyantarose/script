<?php
// /var/www/html/_vps_secret.php

function _vps_read_expected_secret(): string {
  $path = '/etc/niyantarose/ssh_secret.txt';
  $s = @file_get_contents($path);
  if ($s === false) return '';
  return trim(str_replace(["\r", "\n"], '', $s));
}

function _vps_get_header(string $name): string {
  // 1) getallheaders()
  if (function_exists('getallheaders')) {
    $h = getallheaders();
    foreach ($h as $k => $v) {
      if (strcasecmp($k, $name) === 0) return trim((string)$v);
    }
  }
  // 2) $_SERVER fallback (X-VPS-Secret -> HTTP_X_VPS_SECRET)
  $key = 'HTTP_' . strtoupper(str_replace('-', '_', $name));
  if (isset($_SERVER[$key])) return trim((string)$_SERVER[$key]);
  return '';
}

function vps_get_request_secret(): string {
  // ヘッダ優先（GASはこれ）
  $h = _vps_get_header('X-VPS-Secret');
  if ($h !== '') return $h;

  // 後方互換：POST secret も許容（curlや旧実装対策）
  if (isset($_POST['secret'])) return trim((string)$_POST['secret']);

  return '';
}

function vps_json_respond(array $arr, int $code): void {
  http_response_code($code);
  header('Content-Type: application/json; charset=utf-8');
  echo json_encode($arr, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
  exit;
}

function require_secret_or_403(): void {
  $expected = _vps_read_expected_secret();
  if ($expected === '') {
    vps_json_respond(['ok'=>false, 'error'=>'SECRET_FILE_MISSING'], 500);
  }

  $got = vps_get_request_secret();
  if ($got === '' || !hash_equals($expected, $got)) {
    vps_json_respond(['ok'=>false, 'error'=>'FORBIDDEN'], 403);
  }
}
