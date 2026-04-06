<?php
header('Content-Type: text/plain; charset=UTF-8');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
  http_response_code(405);
  echo "POST_ONLY\n";
  exit;
}

// --- SECRET check (required) ---
$server_secret = @file_get_contents('/etc/niyantarose/ssh_secret.txt');
$server_secret = $server_secret ? trim($server_secret) : '';

$client_secret = $_SERVER['HTTP_X_VPS_SECRET'] ?? ($_POST['secret'] ?? '');
$client_secret = is_string($client_secret) ? trim($client_secret) : '';

if ($server_secret === '' || $client_secret === '' || !hash_equals($server_secret, $client_secret)) {
  http_response_code(403);
  echo "FORBIDDEN\n";
  exit;
}

$cmd = $_POST['command'] ?? '';
$cmd = is_string($cmd) ? trim($cmd) : '';

if ($cmd === '') {
  echo "NO_COMMAND\n";
  exit;
}

/**
 * 最低限の防御：
 * - 危険な記号を拒否（完全ではないが大事故を減らす）
 * - 必要なら allow list 方式に変更推奨
 */
if (preg_match('/[;&|`$<>\n\r]/', $cmd)) {
  http_response_code(400);
  echo "REJECTED_CHARS\n";
  exit;
}

// ここを “許可コマンドだけ” に絞るとさらに安全
$allowed = [
  'rclone ', 'ls ', 'du ', 'find ', 'bash ', 'sh ', 'php ', 'python3 ', 'node ',
  'whoami', 'id ', 'pwd ',
];


$ok = false;
foreach ($allowed as $pfx) {
  if (stripos($cmd, $pfx) === 0) { $ok = true; break; }
}
if (!$ok) {
  http_response_code(400);
  echo "NOT_ALLOWED\n";
  exit;
}

$output = [];
$return_var = 0;
exec($cmd . " 2>&1", $output, $return_var);

if ($return_var === 0) {
  echo "SUCCESS\n";
  if (!empty($output)) echo implode("\n", $output);
  echo "\n";
} else {
  echo "ERROR\n";
  if (!empty($output)) echo implode("\n", $output);
  echo "\n";
}
