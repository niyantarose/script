function checkDetailedAuthStatus() {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty('ACCESS_TOKEN');
  const refreshToken = props.getProperty('REFRESH_TOKEN');
  const expiresAt = props.getProperty('TOKEN_EXPIRES_AT');
  
  Logger.log('=== 認証状態詳細 ===');
  Logger.log('ACCESS_TOKEN存在: ' + !!accessToken);
  Logger.log('REFRESH_TOKEN存在: ' + !!refreshToken);
  
  if (expiresAt) {
    const expireDate = new Date(Number(expiresAt));
    const now = new Date();
    Logger.log('トークン有効期限: ' + expireDate.toLocaleString('ja-JP'));
    Logger.log('現在時刻: ' + now.toLocaleString('ja-JP'));
    Logger.log('有効期限まで: ' + Math.round((expireDate.getTime() - now.getTime()) / (1000 * 60 * 60)) + '時間');
  }
  
  // 実際にAPIアクセステスト
  try {
    const token = getValidAccessToken();
    Logger.log('API アクセステスト: 成功');
    return true;
  } catch (e) {
    Logger.log('API アクセステスト: 失敗 - ' + e.message);
    return false;
  }
}