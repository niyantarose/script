function testSimpleEndpoint() {
  const token = PropertiesService.getScriptProperties().getProperty('NOCO_TOKEN');
  
  // よりシンプルなエンドポイントでテスト
  const testUrl = `http://${VPS_IP}:8080/api/v1/db/meta/projects`;
  
  const response = UrlFetchApp.fetch(testUrl, {
    method: 'GET',
    headers: {'xc-auth': token},
    muteHttpExceptions: true
  });
  
  console.log('ステータス:', response.getResponseCode());
  console.log('レスポンス:', response.getContentText());
}