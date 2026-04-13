function createCustomTrigger() {
  // 既存の同一関数のトリガーが残っていたら削除（任意）
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshYahooAccessToken') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // ここで「50分おき」のトリガーを作成
  ScriptApp
    .newTrigger('refreshYahooAccessToken')// ← ここを変えれば好きな関数にする
    .timeBased()
    .everyMinutes(30)    // ← ここを変えれば好きな分数（1～59分）にできます
    .create();
}
