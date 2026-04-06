<?php
header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
  http_response_code(405);
  echo json_encode(['ok' => false, 'error' => 'Method not allowed']);
  exit;
}

$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

if (!$data || empty($data['base64'])) {
  http_response_code(400);
  echo json_encode(['ok' => false, 'error' => 'Invalid JSON or base64 missing']);
  exit;
}

$filename = isset($data['filename']) ? (string)$data['filename'] : 'image.webp';
$bin = base64_decode($data['base64'], true);

if ($bin === false) {
  http_response_code(400);
  echo json_encode(['ok' => false, 'error' => 'base64 decode failed']);
  exit;
}

$tmpWebp = tempnam(sys_get_temp_dir(), 'w2j_') . '.webp';
$tmpJpg  = tempnam(sys_get_temp_dir(), 'w2j_') . '.jpg';
file_put_contents($tmpWebp, $bin);

$ok = false;
$err = '';

try {
  // 1) GD（imagecreatefromwebp）が使えるなら最優先
  if (function_exists('imagecreatefromwebp')) {
    $im = @imagecreatefromwebp($tmpWebp);
    if ($im) {
      $w = imagesx($im);
      $h = imagesy($im);
      $bg = imagecreatetruecolor($w, $h);
      $white = imagecolorallocate($bg, 255, 255, 255);
      imagefill($bg, 0, 0, $white);
      imagecopy($bg, $im, 0, 0, 0, 0, $w, $h);

      $ok = imagejpeg($bg, $tmpJpg, 92);
      imagedestroy($bg);
      imagedestroy($im);
    } else {
      $err = 'GD: imagecreatefromwebp failed';
    }
  }

  // 2) GDがダメなら Imagick を試す
  if (!$ok && class_exists('Imagick')) {
    $im = new Imagick();
    $im->readImage($tmpWebp);
    $im->setImageFormat('jpeg');
    $im->setImageCompressionQuality(92);
    $im->writeImage($tmpJpg);
    $im->clear();
    $im->destroy();
    $ok = file_exists($tmpJpg) && filesize($tmpJpg) > 0;
    if (!$ok) $err = 'Imagick: write failed';
  }

} catch (Throwable $e) {
  $err = $e->getMessage();
}

if (!$ok) {
  @unlink($tmpWebp);
  @unlink($tmpJpg);
  http_response_code(500);
  echo json_encode(['ok' => false, 'error' => $err ?: 'convert failed']);
  exit;
}

$jpgBin = file_get_contents($tmpJpg);
@unlink($tmpWebp);
@unlink($tmpJpg);

echo json_encode([
  'ok' => true,
  'base64' => base64_encode($jpgBin),
  'out' => preg_replace('/\.webp$/i', '.jpg', $filename)
], JSON_UNESCAPED_UNICODE);
