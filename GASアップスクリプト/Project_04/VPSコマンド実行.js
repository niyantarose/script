function VPS_コマンド実行_(cmd) {
  const secret = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');
  if (!secret) throw new Error('ScriptProperties に VPS_SECRET が未設定');

  const url = 'https://img.niyantarose.com/ssh_execute.php';

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'X-VPS-SECRET': secret },
    payload: { command: cmd },
    muteHttpExceptions: true,
  });

  const text = res.getContentText();
  const code = res.getResponseCode();
  if (code !== 200) throw new Error('VPS HTTP ' + code + '\n' + text);

  if (!text.startsWith('SUCCESS')) {
    throw new Error('VPS ERROR\n' + text);
  }
  return text;
}

function checkMyProperty() {
  const secret = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');

  if (!secret) {
    Logger.log('VPS_SECRET は未設定（null/空）');
    return;
  }
  Logger.log('VPS_SECRET は設定済み（長さ=' + secret.length + '）');
}

function fixMyProperty() {

  // 登録したい名前と値をここで指定

  const key = 'VPS_SECRET';

  // ↓ここに、登録したい長い文字列をコピペしてください

  const val = 'bce535b6993c5eda1ddf8eb9ff454dcc41ee142b354571df32b5f62515487dbf'; 



  // 1. 強制的に登録（上書き）

  PropertiesService.getScriptProperties().setProperty(key, val);



  // 2. その場ですぐに確認

  const check = PropertiesService.getScriptProperties().getProperty(key);

  

  console.log('▼▼▼ 結果 ▼▼▼');

  console.log('登録された値: ' + check);

}

function VPS疎通テスト_echo() {
  const url = 'https://img.niyantarose.com/ssh_execute.php';
  const secret = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      // ここはサーバ側の仕様に合わせる（もし secret チェックがあるなら渡す）
      secret: secret,
      command: 'whoami && pwd && date'
    },
    muteHttpExceptions: true
  });

  Logger.log('HTTP=' + res.getResponseCode());
  Logger.log(res.getContentText());
}


function DEBUG_スクリプトプロパティ容量チェック() {
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();

  const arr = Object.entries(props).map(([k, v]) => {
    const keyBytes = Utilities.newBlob(String(k)).getBytes().length;
    const valBytes = Utilities.newBlob(String(v ?? '')).getBytes().length;
    return { key: k, bytes: keyBytes + valBytes, valBytes };
  }).sort((a, b) => b.bytes - a.bytes);

  const total = arr.reduce((s, x) => s + x.bytes, 0);

  Logger.log("keys=" + arr.length);
  Logger.log("total_bytes=" + total + " (limit ~512000 bytes)");
  Logger.log("top10=" + JSON.stringify(arr.slice(0, 10), null, 2));
}

function _プロパティ強制保存() {
  const p = PropertiesService.getScriptProperties();
  p.setProperties({
    Y_SCOPE: 'openid',
    Y_SELLER_ID: 'niyantarose',
    // VPS_SECRET: '（ここに秘密文字列）'
  }, true);

  Logger.log('OK');
}

function VPS_コマンド実行_(cmd) {
  const url = 'https://img.niyantarose.com/ssh_execute.php';
  const secret = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'X-VPS-SECRET': secret },
    payload: { command: cmd },
    muteHttpExceptions: true,
    followRedirects: true,
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code !== 200) throw new Error('HTTP=' + code + '\n' + text);
  if (!String(text).startsWith('SUCCESS')) throw new Error(text);

  return text; // SUCCESS\n...
}

// テスト
function VPS疎通テスト2() {
  const out = VPS_コマンド実行_('bash -lc whoami');
  console.log(out);
}

function testImageApi() {
  const url = 'https://img.niyantarose.com/yahoo_image_api.php';

  const payload = {
    item_code: 'TEST-001',
    images: [
      'https://via.placeholder.com/800.jpg',
      'https://via.placeholder.com/600.jpg'
    ]
  };

  const res = fetchYahooApi(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
  });

  Logger.log(res);
}

function testImageApi() {
  const url = 'https://img.niyantarose.com/yahoo_image_api.php';

  const payload = {
    item_code: 'TEST-001',
    images: [
      'https://via.placeholder.com/800.jpg',
      'https://via.placeholder.com/600.jpg'
    ]
  };

  const res = fetchYahooApi(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
  });

  Logger.log(res);
}

// ↓↓↓ この関数を1回だけ実行してください ↓↓↓
function setNewSecret() {
  const KEY = 'IMGCV_VPS_SECRET'; // v29コードの設定名に合わせます
  const NEW_VAL = 'f880e5857bee500a43efa450f92419ae498d3fbf14f4959ea32ebdf5785b3870';

  // プロパティを強制上書き
  PropertiesService.getScriptProperties().setProperty(KEY, NEW_VAL);

  // 確認ログ
  const current = PropertiesService.getScriptProperties().getProperty(KEY);
  if (current === NEW_VAL) {
    console.log('✅ 更新完了: IMGCV_VPS_SECRET は新しい値になりました。');
    console.log('現在値: ' + current.slice(0, 10) + '...'); 
  } else {
    console.error('❌ 更新失敗: 何かがおかしいです。');
  }
}
function checkSecret() {
  const v = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');
  console.log('VPS_SECRET length=' + (v ? v.length : 'null'));
  console.log('VPS_SECRET head=' + (v ? v.substring(0, 8) : 'null'));
}

function forceUpdateSecret() {
  PropertiesService.getScriptProperties().setProperty(
    'VPS_SECRET',
    'f880e5857bee500a43efa450f92419ae498d3fbf14f4959ea32ebdf5785b3870'
  );
  // 確認
  const v = PropertiesService.getScriptProperties().getProperty('VPS_SECRET');
  console.log('Updated: head=' + v.substring(0, 8) + ' len=' + v.length);
}

function セットプロパティ() {
  // プロパティサービスを呼び出す
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // 保存したいキーと値を指定
  const key = 'VPS_SC';
  const value = 'bce535b6993c5eda1ddf8eb9ff454dcc41ee142b354571df32b5f62515487dbf';
  
  // プロパティを設定
  scriptProperties.setProperty(key, value);
  
  console.log('プロパティ ' + key + ' の設定が完了しました。');
}

function CDNアクセステスト() {
  const sh = SpreadsheetApp.getActive().getSheetByName('①商品入力シート');
  const url = sh.getRange(3, 19).getValue(); // 3行目S列のCDN URL
  console.log('テストURL: ' + url);
  
  const res = UrlFetchApp.fetch(url, {
    headers: { 
      'User-Agent': 'Mozilla/5.0',
      'Referer': 'https://shopping.yahoo.co.jp/'
    },
    muteHttpExceptions: true
  });
  
  console.log('HTTPステータス: ' + res.getResponseCode());
  console.log('Content-Type: ' + res.getHeaders()['Content-Type']);
  console.log('サイズ: ' + res.getContent().length + ' bytes');
  const head = res.getContentText().substring(0, 10);
  console.log('先頭バイト: ' + head);
}

function VPSサーバーURL確認() {
  const sh = SpreadsheetApp.getActive().getSheetByName('①商品入力シート');
  const url = sh.getRange(3, 19).getValue();
  console.log('S列URL: ' + url);
  
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  console.log('HTTP: ' + res.getResponseCode());
  console.log('ContentType: ' + res.getHeaders()['Content-Type']);
  console.log('サイズ: ' + res.getContent().length);
}