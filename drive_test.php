<?php
require __DIR__ . '/vendor/autoload.php';

$serviceAccountPath = '/home/ubuntu/secret/service-account.json';

// 商品画像ルート or 免責ルート、どっちでもOK（まずは商品画像ルート推奨）
$rootFolderId = '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz';

$client = new Google\Client();
$client->setAuthConfig($serviceAccountPath);
$client->addScope(Google\Service\Drive::DRIVE_READONLY);

$drive = new Google\Service\Drive($client);

header('Content-Type: text/plain; charset=utf-8');

try {
  // 共有ドライブ対応：まずメタ取得（driveIdが取れると検索が安定）
  $meta = $drive->files->get($rootFolderId, [
    'fields' => 'id,name,driveId',
    'supportsAllDrives' => true
  ]);
  $driveId = $meta->driveId ?? null;

  // 直下を少し列挙してみる
  $opt = [
    'q' => sprintf("'%s' in parents and trashed = false", $rootFolderId),
    'fields' => 'files(id,name,mimeType,size)',
    'pageSize' => 20,
    'supportsAllDrives' => true,
    'includeItemsFromAllDrives' => true,
  ];
  if ($driveId) {
    $opt['corpora'] = 'drive';
    $opt['driveId'] = $driveId;
  }

  $list = $drive->files->listFiles($opt);
  $files = $list->getFiles();

  echo "root: {$meta->name} ({$meta->id}) driveId=" . ($driveId ?: '(none)') . PHP_EOL;
  echo "items:" . PHP_EOL;
  foreach ($files as $f) {
    echo "- {$f->getName()}  {$f->getId()}  {$f->getMimeType()}  size={$f->getSize()}" . PHP_EOL;
  }

} catch (Exception $e) {
  echo "ERROR: " . $e->getMessage() . PHP_EOL;
}
