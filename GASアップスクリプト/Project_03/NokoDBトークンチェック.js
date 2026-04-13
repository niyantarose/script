function testToken() {
  const token = PropertiesService.getScriptProperties().getProperty('NOCO_TOKEN');
  console.log('現在のトークン:', token);
  
  // 簡単な接続テスト
  const response = UrlFetchApp.fetch(NOCO_ENDPOINT, {
    method: 'GET',
    headers: {'xc-auth': token},
    muteHttpExceptions: true
  });
  
  console.log('ステータスコード:', response.getResponseCode());
  console.log('レスポンス:', response.getContentText());
}