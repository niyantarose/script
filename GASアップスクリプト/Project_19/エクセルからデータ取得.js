// ====== 設定エリア ======
const TENANT_ID = '234dd4aa-6f7b-44b1-b4f0-b54e2beeddad';
const CLIENT_ID = 'f2137667-e77d-49fb-9e75-2c8cf78ba199';
const CLIENT_SECRET_PROPERTY = 'MS_GRAPH_CLIENT_SECRET';
const DRIVE_ITEM_ID = '42168735-D1F5-4DC6-A9DA-24503D028A40';
const USER_EMAIL = 'niyantarose@logosnochikara.onmicrosoft.com';
const SHEET_NAMES = ['発注リスト', 'EMS'];
const LAST_COL = 'AA';
const LAST_COL_NUM = 27; // AA = 27列目

function getClientSecret_() {
  const secret = PropertiesService.getScriptProperties().getProperty(CLIENT_SECRET_PROPERTY);
  if (!secret) {
    throw new Error(`Script Properties に ${CLIENT_SECRET_PROPERTY} がありません。メニュー「Excel同期」→「初期設定：CLIENT_SECRETを保存」から設定してください。`);
  }
  return secret;
}

// ====== メニュー（開いたとき「Excel同期」メニューが出る）======
function Excel同期メニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('Excel同期')
    .addItem('全部実行（Excel取込＋差分＋カレンダー）', 'syncAll')
    .addItem('Excel取込のみ', 'importSharePointExcel')
    .addItem('差分反映のみ', 'syncDifferences')
    .addItem('EMSカレンダー更新のみ', 'syncEmsCalendar')
    .addItem('EMSリスト：重複行を確認して削除', 'EMSリスト_重複行を確認して削除')
    .addItem('EMSリスト：下に混ざった古い行を確認して削除', 'EMSリスト_下に混ざった古い行を確認して削除')
    .addItem('EMSリスト：空欄行を同期データで補完', 'EMSリスト_空欄行をEMS同期データから補完')
    .addSeparator()
    .addItem('商品マスタ：足りないデータだけ発注から補完', '商品マスタ_足りないデータだけ発注から補完')
    .addItem('商品マスタ：既存データを発注から更新（重さ等）', '商品マスタ_既存データを発注から補完更新')
    .addItem('商品マスタ：重複行を確認して削除', '商品マスタ_重複行を確認して削除')
    .addSeparator()
    .addItem('初期設定：CLIENT_SECRETを保存', 'MS_GRAPH_CLIENT_SECRETを設定')
    .addItem('初期設定：CLIENT_SECRETを確認', 'MS_GRAPH_CLIENT_SECRETを確認')
    .addToUi();
}

function MS_GRAPH_CLIENT_SECRETを設定() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Microsoft Graph CLIENT_SECRET 設定',
    'Azureのクライアントシークレット値を貼り付けてOKを押してください。入力欄には表示されるので周囲に注意してください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    SpreadsheetApp.getActive().toast('キャンセルしました', 'Excel同期', 3);
    return;
  }

  const secret = String(res.getResponseText() || '').trim();
  if (!secret) {
    ui.alert('空欄なので保存していません。');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(CLIENT_SECRET_PROPERTY, secret);
  SpreadsheetApp.getActive().toast('CLIENT_SECRETをScript Propertiesに保存しました。', 'Excel同期', 5);
}

function MS_GRAPH_CLIENT_SECRETを確認() {
  const exists = !!PropertiesService.getScriptProperties().getProperty(CLIENT_SECRET_PROPERTY);
  SpreadsheetApp.getUi().alert(
    exists
      ? `${CLIENT_SECRET_PROPERTY} は設定済みです。`
      : `${CLIENT_SECRET_PROPERTY} は未設定です。メニュー「Excel同期」→「初期設定：CLIENT_SECRETを保存」から設定してください。`
  );
}
// ====== 自動実行の設定（最初に1回だけ手動実行する）======
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'importSharePointExcel' || f === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyDays(1).atHour(9).nearMinute(0).create();
  ScriptApp.newTrigger('syncAll').timeBased().everyDays(1).atHour(15).nearMinute(0).create();
  Logger.log('朝9時・夕方15時に syncAll を実行するトリガーを設定しました');
}

// ====== Graph API呼び出し（リトライ付き）======
function graphGet_(url, token, sessionId) {
  const options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (sessionId) options.headers['Workbook-Session-Id'] = sessionId;

  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    const res = UrlFetchApp.fetch(url, options);
    let json;
    try { json = JSON.parse(res.getContentText()); } catch (e) {
      lastErr = 'JSON解析失敗 (HTTP ' + res.getResponseCode() + ')';
      Utilities.sleep(2000 * attempt);
      continue;
    }
    if (!json.error) return json;
    // UnknownError / 5xx系は待ってリトライ
    const code = json.error.code || '';
    if (code === 'UnknownError' || res.getResponseCode() >= 500 || code === 'ServiceUnavailable' || code === 'TooManyRequests') {
      lastErr = JSON.stringify(json.error);
      Logger.log(`一時エラー（${attempt}回目）リトライします: ${code}`);
      Utilities.sleep(3000 * attempt);
      continue;
    }
    return json; // リトライ不能なエラーはそのまま返す
  }
  return { error: { code: 'RetryExhausted', message: lastErr } };
}

// ====== ワークブックセッションの作成・終了 ======
function createWorkbookSession_(token) {
  const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${DRIVE_ITEM_ID}/workbook/createSession`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ persistChanges: false }), // 読み取り専用
    muteHttpExceptions: true
  });
  try {
    const json = JSON.parse(res.getContentText());
    return json.id || null;
  } catch (e) { return null; }
}

function closeWorkbookSession_(token, sessionId) {
  if (!sessionId) return;
  try {
    UrlFetchApp.fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${DRIVE_ITEM_ID}/workbook/closeSession`, {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Workbook-Session-Id': sessionId },
        muteHttpExceptions: true
      });
  } catch (e) { /* 失敗しても無視 */ }
}

// ====== メイン処理 ======
function importSharePointExcel() {
  const t0 = Date.now();
  const lap = label => Logger.log(`[${((Date.now() - t0) / 1000).toFixed(1)}秒] ${label}`);

  const token = getMicrosoftAccessToken();
  if (!token) return;
  lap('トークン取得');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let xlsxParts = null;
  try {
    xlsxParts = fetchXlsxParts_(token);
    lap('xlsxダウンロード＆展開');
  } catch (e) {
    Logger.log('背景色の取得に失敗したので色なしで続行: ' + e.message);
  }

  // セッションを作って全リクエストで使い回す（安定性対策）
  const sessionId = createWorkbookSession_(token);
  lap('セッション作成' + (sessionId ? '' : '（失敗・セッションなしで続行）'));

  const baseUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${DRIVE_ITEM_ID}/workbook/worksheets`;

  try {
    SHEET_NAMES.forEach(sheetName => {
      const enc = encodeURIComponent(sheetName);

      // ① 実データの最終行
      const meta = graphGet_(
        `${baseUrl}('${enc}')/usedRange(valuesOnly=true)?$select=address`, token, sessionId);
      if (meta.error) {
        Logger.log(`${sheetName} メタ取得エラー: ` + JSON.stringify(meta.error)); return;
      }
      const m = String(meta.address).match(/(\d+)\s*$/);
      if (!m) return;
      const lastRow = parseInt(m[1], 10);
      lap(`${sheetName} 最終行取得 (${lastRow}行)`);

      // ② 値と表示形式（リトライ付き）
      const json = graphGet_(
        `${baseUrl}('${enc}')/range(address='A1:${LAST_COL}${lastRow}')?$select=values,numberFormat`,
        token, sessionId);
      if (json.error || !json.values || json.values.length === 0) {
        Logger.log(`${sheetName} データ取得エラー: ` + JSON.stringify(json.error || '空')); return;
      }
      lap(`${sheetName} 値・表示形式取得`);

      const data = json.values;
      const targetSheetName = sheetName + '同期データ';
      let targetSheet = ss.getSheetByName(targetSheetName);
      const isNewSheet = !targetSheet;
      if (isNewSheet) targetSheet = ss.insertSheet(targetSheetName);

      targetSheet.clear();
      const range = targetSheet.getRange(1, 1, data.length, data[0].length);
      range.setValues(data);

      if (json.numberFormat && json.numberFormat.length === data.length) {
        range.setNumberFormats(json.numberFormat.map(r => r.map(sanitizeNumberFormat_)));
      }
      lap(`${sheetName} 書き込み＆表示形式`);

      if (xlsxParts && xlsxParts.sheetXml[sheetName]) {
        const grid = parseSheetBackgroundsFast_(
          xlsxParts.sheetXml[sheetName], xlsxParts.styleFills, data.length);
        range.setBackgrounds(cropBackgrounds_(grid, data.length, data[0].length));
        lap(`${sheetName} 背景色適用`);
      }

      formatSyncedSheet_(targetSheet, sheetName, data.length, data[0].length, isNewSheet);
      lap(`${sheetName} 完了`);
    });
  } finally {
    closeWorkbookSession_(token, sessionId);
  }

  lap('全処理完了');
}

// ====== xlsxをDLして必要な部品（スタイル表・シートXML文字列）だけ取り出す ======
function fetchXlsxParts_(token) {
  const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${DRIVE_ITEM_ID}/content`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('xlsxダウンロード失敗: HTTP ' + res.getResponseCode());
  }
  const files = Utilities.unzip(res.getBlob().setContentType('application/zip'));
  const fileMap = {};
  files.forEach(f => { fileMap[f.getName()] = f; });

  // テーマ色（正規表現で抽出）
  const themeColors = parseThemeColorsFast_(
    fileMap['xl/theme/theme1.xml'] ? fileMap['xl/theme/theme1.xml'].getDataAsString() : null);

  // スタイル番号 → 塗りつぶし色
  const styleFills = parseStyleFillsFast_(fileMap['xl/styles.xml'].getDataAsString(), themeColors);

  // シート名 → XML文字列
  const wbXml = fileMap['xl/workbook.xml'].getDataAsString();
  const relsXml = fileMap['xl/_rels/workbook.xml.rels'].getDataAsString();
  const relMap = {};
  let rm;
  const relRe = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  while ((rm = relRe.exec(relsXml)) !== null) relMap[rm[1]] = rm[2];

  const sheetXml = {};
  const sheetRe = /<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g;
  let sm;
  while ((sm = sheetRe.exec(wbXml)) !== null) {
    const name = decodeXmlEntities_(sm[1]);
    if (SHEET_NAMES.indexOf(name) === -1) continue;
    let target = relMap[sm[2]];
    if (!target) continue;
    target = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target;
    if (fileMap[target]) sheetXml[name] = fileMap[target].getDataAsString();
  }

  return { styleFills: styleFills, sheetXml: sheetXml };
}

// ====== styles.xml を正規表現で解析 ======
function parseStyleFillsFast_(stylesXml, themeColors) {
  // <fills>～</fills> から各fillの色を順番に取り出す
  const fills = [];
  const fillsBlock = (stylesXml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/) || [])[1] || '';
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(fillsBlock)) !== null) {
    const body = fm[1];
    let color = null;
    if (/patternType="solid"/.test(body)) {
      const fg = (body.match(/<fgColor([^/>]*)\/?>/) || [])[1] || '';
      color = resolveColorFast_(fg, themeColors);
    }
    fills.push(color);
  }

  // <cellXfs>～</cellXfs> の各xfが参照するfillIdを色に変換
  const styleFills = [];
  const xfsBlock = (stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
  const xfRe = /<xf\b[^>]*>/g;
  let xm;
  while ((xm = xfRe.exec(xfsBlock)) !== null) {
    const idm = xm[0].match(/fillId="(\d+)"/);
    styleFills.push(idm ? (fills[parseInt(idm[1], 10)] || null) : null);
  }
  return styleFills;
}

// fgColorタグの属性文字列から色を決定
function resolveColorFast_(attrs, themeColors) {
  let m = attrs.match(/rgb="([0-9A-Fa-f]+)"/);
  if (m) return '#' + m[1].slice(-6);

  m = attrs.match(/theme="(\d+)"/);
  if (m && themeColors) {
    let hex = themeColors[parseInt(m[1], 10)];
    if (hex) {
      const tm = attrs.match(/tint="(-?[\d.]+)"/);
      if (tm) hex = applyTint_(hex, parseFloat(tm[1]));
      return '#' + hex;
    }
  }

  m = attrs.match(/indexed="(\d+)"/);
  if (m) {
    const hex = INDEXED_COLORS_[parseInt(m[1], 10)];
    return hex ? '#' + hex : null;
  }
  return null;
}

// ====== theme1.xml を正規表現で解析 ======
function parseThemeColorsFast_(themeXml) {
  if (!themeXml) return null;
  const colorOf = name => {
    const block = (themeXml.match(new RegExp('<a:' + name + '>([\\s\\S]*?)</a:' + name + '>')) || [])[1] || '';
    let m = block.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    if (m) return m[1];
    m = block.match(/lastClr="([0-9A-Fa-f]{6})"/);
    if (m) return m[1];
    return 'FFFFFF';
  };
  return [colorOf('lt1'), colorOf('dk1'), colorOf('lt2'), colorOf('dk2'),
          colorOf('accent1'), colorOf('accent2'), colorOf('accent3'),
          colorOf('accent4'), colorOf('accent5'), colorOf('accent6'),
          colorOf('hlink'), colorOf('folHlink')];
}

// ====== シートXMLからセル背景色を正規表現で抽出（rowLimitまで）======
function parseSheetBackgroundsFast_(sheetXmlStr, styleFills, rowLimit) {
  const grid = {};
  // セルは <c r="D13" s="5" t="s"> か <c r="D13" s="5"/> の形
  const cellRe = /<c\b([^>]*)>/g;
  let cm;
  while ((cm = cellRe.exec(sheetXmlStr)) !== null) {
    const attrs = cm[1];
    const sM = attrs.match(/\bs="(\d+)"/);
    if (!sM) continue;
    const color = styleFills[parseInt(sM[1], 10)];
    if (!color) continue;
    const rM = attrs.match(/\br="([A-Z]+)(\d+)"/);
    if (!rM) continue;
    const row = parseInt(rM[2], 10);
    if (row > rowLimit) continue;
    let col = 0;
    for (let i = 0; i < rM[1].length; i++) col = col * 26 + (rM[1].charCodeAt(i) - 64);
    if (col > LAST_COL_NUM) continue;
    grid[row + ',' + col] = color;
  }
  return grid;
}

function cropBackgrounds_(grid, rows, cols) {
  const out = [];
  for (let r = 1; r <= rows; r++) {
    const rowArr = [];
    for (let c = 1; c <= cols; c++) rowArr.push(grid[r + ',' + c] || null);
    out.push(rowArr);
  }
  return out;
}

function decodeXmlEntities_(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

// ====== xlsxを解析して背景色を取り出す ======
function fetchXlsxBackgrounds_(token) {
  const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${DRIVE_ITEM_ID}/content`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('xlsxダウンロード失敗: HTTP ' + res.getResponseCode());
  }
  const files = Utilities.unzip(res.getBlob().setContentType('application/zip'));
  const fileMap = {};
  files.forEach(f => { fileMap[f.getName()] = f; });

  const nsMain = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  const nsR = XmlService.getNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  const nsPkg = XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships');

  const themeColors = parseThemeColors_(fileMap['xl/theme/theme1.xml']);
  const styleFills = parseStyleFills_(fileMap['xl/styles.xml'], nsMain, themeColors);

  // シート名 → sheetN.xml の対応を作る
  const wb = XmlService.parse(fileMap['xl/workbook.xml'].getDataAsString()).getRootElement();
  const rels = XmlService.parse(fileMap['xl/_rels/workbook.xml.rels'].getDataAsString()).getRootElement();
  const relMap = {};
  rels.getChildren('Relationship', nsPkg).forEach(r => {
    relMap[r.getAttribute('Id').getValue()] = r.getAttribute('Target').getValue();
  });

  const result = {};
  wb.getChild('sheets', nsMain).getChildren('sheet', nsMain).forEach(sh => {
    const name = sh.getAttribute('name').getValue();
    if (SHEET_NAMES.indexOf(name) === -1) return;
    let target = relMap[sh.getAttribute('id', nsR).getValue()];
    target = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target;
    if (fileMap[target]) {
      result[name] = parseSheetBackgrounds_(fileMap[target], nsMain, styleFills);
    }
  });
  return result;
}

function parseStyleFills_(stylesFile, nsMain, themeColors) {
  const root = XmlService.parse(stylesFile.getDataAsString()).getRootElement();
  const fills = [];
  const fillsEl = root.getChild('fills', nsMain);
  if (fillsEl) {
    fillsEl.getChildren('fill', nsMain).forEach(f => {
      let color = null;
      const p = f.getChild('patternFill', nsMain);
      if (p) {
        const typeAttr = p.getAttribute('patternType');
        if (typeAttr && typeAttr.getValue() === 'solid') {
          color = resolveColor_(p.getChild('fgColor', nsMain), themeColors);
        }
      }
      fills.push(color);
    });
  }
  const styleFills = [];
  const xfsEl = root.getChild('cellXfs', nsMain);
  if (xfsEl) {
    xfsEl.getChildren('xf', nsMain).forEach(xf => {
      const a = xf.getAttribute('fillId');
      styleFills.push(fills[a ? parseInt(a.getValue(), 10) : 0] || null);
    });
  }
  return styleFills;
}

function resolveColor_(colorEl, themeColors) {
  if (!colorEl) return null;
  const rgb = colorEl.getAttribute('rgb');
  if (rgb) return '#' + rgb.getValue().slice(-6);
  const theme = colorEl.getAttribute('theme');
  if (theme && themeColors) {
    let hex = themeColors[parseInt(theme.getValue(), 10)];
    if (hex) {
      const tint = colorEl.getAttribute('tint');
      if (tint) hex = applyTint_(hex, parseFloat(tint.getValue()));
      return '#' + hex;
    }
  }
  const indexed = colorEl.getAttribute('indexed');
  if (indexed) {
    const hex = INDEXED_COLORS_[parseInt(indexed.getValue(), 10)];
    return hex ? '#' + hex : null;
  }
  return null;
}

function parseThemeColors_(themeFile) {
  if (!themeFile) return null;
  const nsA = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
  const root = XmlService.parse(themeFile.getDataAsString()).getRootElement();
  const scheme = root.getChild('themeElements', nsA).getChild('clrScheme', nsA);
  const colorOf = name => {
    const el = scheme.getChild(name, nsA);
    if (!el) return 'FFFFFF';
    const srgb = el.getChild('srgbClr', nsA);
    if (srgb) return srgb.getAttribute('val').getValue();
    const sys = el.getChild('sysClr', nsA);
    if (sys && sys.getAttribute('lastClr')) return sys.getAttribute('lastClr').getValue();
    return 'FFFFFF';
  };
  // theme属性のインデックス順（lt1とdk1、lt2とdk2が入れ替わるのがExcelの仕様）
  return [colorOf('lt1'), colorOf('dk1'), colorOf('lt2'), colorOf('dk2'),
          colorOf('accent1'), colorOf('accent2'), colorOf('accent3'),
          colorOf('accent4'), colorOf('accent5'), colorOf('accent6'),
          colorOf('hlink'), colorOf('folHlink')];
}

function applyTint_(hex, tint) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const v = parseInt(hex.substr(i, 2), 16);
    let nv = tint < 0 ? v * (1 + tint) : v + (255 - v) * tint;
    nv = Math.max(0, Math.min(255, Math.round(nv)));
    out += ('0' + nv.toString(16)).slice(-2);
  }
  return out;
}

const INDEXED_COLORS_ = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  '800000','008000','000080','808000','800080','008080','C0C0C0','808080',
  '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
  '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF',
  '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
  '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696',
  '003366','339966','003300','333300','993300','993366','333399','333333'
];

function parseSheetBackgrounds_(sheetFile, nsMain, styleFills) {
  const root = XmlService.parse(sheetFile.getDataAsString()).getRootElement();
  const sheetData = root.getChild('sheetData', nsMain);
  const grid = {};
  if (sheetData) {
    sheetData.getChildren('row', nsMain).forEach(rowEl => {
      rowEl.getChildren('c', nsMain).forEach(cEl => {
        const s = cEl.getAttribute('s');
        if (!s) return;
        const color = styleFills[parseInt(s.getValue(), 10)];
        if (!color) return;
        const ref = cEl.getAttribute('r').getValue();
        const rm = ref.match(/^([A-Z]+)(\d+)$/);
        if (!rm) return;
        let col = 0;
        for (let i = 0; i < rm[1].length; i++) col = col * 26 + (rm[1].charCodeAt(i) - 64);
        if (col > LAST_COL_NUM) return;
        grid[rm[2] + ',' + col] = color;
      });
    });
  }
  return grid;
}

function cropBackgrounds_(grid, rows, cols) {
  const out = [];
  for (let r = 1; r <= rows; r++) {
    const rowArr = [];
    for (let c = 1; c <= cols; c++) rowArr.push(grid[r + ',' + c] || null);
    out.push(rowArr);
  }
  return out;
}

// ====== 表示形式の調整 ======
function sanitizeNumberFormat_(fmt) {
  if (!fmt || fmt === 'General') return 'General';
  let f = String(fmt);
  f = f.replace(/\[\$-[^\]]+\]/g, '');
  f = f.replace(/;@$/, '');
  f = f.replace(/\*./g, '');
  f = f.trim();

  // Excelの「短い日付形式（ロケール依存）」はm/d/yyyyで返ってくるので日本式に直す
  if (f === 'm/d/yyyy' || f === 'm/d/yy' ||
      f === 'mm/dd/yyyy' || f === 'mm/dd/yy') {
    f = 'yyyy/m/d';
  }

  return f || 'General';
}

// ====== 仕上げの書式（色は本物が来るので最小限に）======
function formatSyncedSheet_(sh, sourceSheetName, rowCount, colCount, isNewSheet) {
  if (!sh || rowCount <= 0 || colCount <= 0) return;

  sh.getRange(1, 1, rowCount, colCount)
    .setFontSize(15)
    .setVerticalAlignment('middle')
    .setWrap(false)
    .setBorder(true, true, true, true, true, true,
      '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  if (sourceSheetName === '発注リスト') sh.setFrozenRows(5);
  if (sourceSheetName === 'EMS') sh.setFrozenRows(2);

  // 行の高さ
  sh.setRowHeights(1, rowCount, 28);

  // ★EMS同期データは3行目のデータを基準に列幅調整
  if (sourceSheetName === 'EMS') {
    fitColumnsBySampleRow_(sh, 3, colCount);
  } else {
    // 発注リスト側は必要ならここも基準行を変えられる
    sh.autoResizeColumns(1, colCount);
  }
}

/// 指定した行の表示文字数を基準に列幅を調整する
function fitColumnsBySampleRow_(sh, sampleRow, colCount) {
  const vals = sh.getRange(sampleRow, 1, 1, colCount).getDisplayValues()[0];

  for (let c = 1; c <= colCount; c++) {
    const text = String(vals[c - 1] || '');

    let width = estimateColumnWidthFromText_(text);

    // ★全体的に少し余裕を持たせる
    width = Math.ceil(width * 1.25);

    // 最小・最大幅
    width = Math.max(70, width);
    width = Math.min(650, width);

    sh.setColumnWidth(c, width);
  }

  // ★EMS同期データで特に見たい列は少し固定気味に広げる
  // D: Tracking # / E: 業者 / F: Description / H: ItemCode
  if (colCount >= 8) {
    sh.setColumnWidth(4, 190);  // D Tracking #
    sh.setColumnWidth(5, 130);  // E 業者
    sh.setColumnWidth(6, 420);  // F Description / Title
    sh.setColumnWidth(8, 220);  // H ItemCode
  }

  // J: Qty / K: Type / M: Unit Price / N: Amount / P: Postage
  if (colCount >= 16) {
    sh.setColumnWidth(10, 80);   // J Qty
    sh.setColumnWidth(11, 150);  // K Type
    sh.setColumnWidth(13, 120);  // M Unit Price
    sh.setColumnWidth(14, 140);  // N Amount
    sh.setColumnWidth(16, 140);  // P Postage
  }
}

// 日本語・韓国語・中国語は広め、英数字も少し余裕ありで計算
function estimateColumnWidthFromText_(text) {
  let units = 0;

  for (const ch of text) {
    if (/[ぁ-んァ-ン一-龥가-힣\u3000-\u303F]/.test(ch)) {
      units += 2.2;
    } else if (/[A-ZMW]/.test(ch)) {
      units += 1.5;
    } else if (/[ilI1.,]/.test(ch)) {
      units += 0.8;
    } else {
      units += 1.15;
    }
  }

  // フォント15pt想定。前より大きめ。
  return Math.ceil(units * 9 + 40);
}

// 日本語・韓国語・中国語は広め、英数字は普通でざっくり幅計算
function estimateColumnWidthFromText_(text) {
  let units = 0;

  for (const ch of text) {
    if (/[ぁ-んァ-ン一-龥가-힣\u3000-\u303F]/.test(ch)) {
      units += 2.0;       // 日本語・韓国語など
    } else if (/[A-ZMW]/.test(ch)) {
      units += 1.3;       // 幅広英字
    } else if (/[ilI1.,]/.test(ch)) {
      units += 0.6;       // 細い文字
    } else {
      units += 1.0;       // その他英数字
    }
  }

  // フォント15pt想定
  return Math.ceil(units * 8 + 28);
}

// ====== Microsoftトークン取得（変更なし）======
function getMicrosoftAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const payload = {
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: getClientSecret_(),
    grant_type: 'client_credentials'
  };
  const response = UrlFetchApp.fetch(url, {
    method: 'post', payload: payload, muteHttpExceptions: true
  });
  const json = JSON.parse(response.getContentText());
  if (json.access_token) return json.access_token;
  Logger.log('トークン取得エラー: ' + response.getContentText());
  return null;
}

// ====== 同期＋差分反映を一括実行 ======
function syncAll() {
  importSharePointExcel();      // SharePoint Excel → 同期データ
  syncDifferences();            // 同期データ → 発注 / EMSリスト
  syncEmsCalendar();            // EMSリスト → Googleカレンダー「EMS追跡」
  buildEmsCalendarSheet();      // EMSリスト → EMSカレンダーシート（帯表示）
}

// ====== 差分反映（同期データ → 発注 / EMSリスト）======
function syncDifferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const r1 = syncDiffToHatchu_(ss);
  const n3 = updateArrivalsInHatchu_(ss);
  const nEmsFill = EMSリスト_空欄行をEMS同期データから補完(true);
  const n4 = updateEmsListFromSync_(ss);   // ★EMSリスト既存行の更新
  const r2 = syncDiffToEmsList_(ss);

  SpreadsheetApp.flush();

  const hachu = ss.getSheetByName('発注');

  // 発注シートの後処理
  if (hachu) {
    if (r1.added > 0 || n3 > 0) {
      runHatchuPostProcess_(hachu, r1);
    }

    // ★商品マスタ：①新規追加 → ②既存値の更新（重さ等） → ③空欄だけ補完
    if (typeof マスタ自動補完_ === 'function') {
      const addedMaster = マスタ自動補完_(true);
      Logger.log(`商品マスタ自動補完：${addedMaster}件追加`);
    }
    if (typeof 商品マスタ_既存データを発注から補完更新 === 'function') {
      商品マスタ_既存データを発注から補完更新(true);   // silent
    }
    if (typeof 商品マスタ_足りないデータだけ発注から補完 === 'function') {
      商品マスタ_足りないデータだけ発注から補完(true); // silent
    }
  }

  // EMSリストの後処理
  const emsSheet = ss.getSheetByName('EMSリスト');
  if (emsSheet && (r2.added > 0 || n4 > 0 || nEmsFill > 0)) {
    runEmsPostProcess_(emsSheet, r2);

    if (hachu && typeof colorKeshikomiAllRows_ === 'function') {
      SpreadsheetApp.flush();
      colorKeshikomiAllRows_(hachu);
    }
  }
EMSリスト_購入No自動補完(true);
  Logger.log(
    `差分反映完了: 発注 ${r1.added}件追加 / EMSリスト ${r2.added}件追加 / 入荷情報 ${n3}セル更新 / EMSリスト更新 ${n4}セル`
      + ` / EMS空欄補完 ${nEmsFill}セル`
  );
}
// ---- 正規化ヘルパー ----
function normOrderNo_(v) {
  let s = String(v || '').trim();
  if (!s) return '';
  const us = s.indexOf('_');
  let head = us >= 0 ? s.slice(0, us) : s;
  const tail = us >= 0 ? s.slice(us) : '';
  head = head.replace(/\D/g, '');           // Wata260402 → 260402
  return head.slice(-6) + tail;             // 20260402 → 260402
}
function normCode_(v) {
  return String(v || '').trim().toUpperCase().replace(/_/g, '-');
}
function normText_(v) {
  return String(v || '').replace(/[\s\u3000]+/g, ' ').trim();
}
function normTrack_(v) {
  return String(v || '').replace(/[\s\u3000]/g, '').toUpperCase();
}

// ---- 発注リスト同期データ → 発注 ----
function syncDiffToHatchu_(ss) {
  const src = ss.getSheetByName('発注リスト同期データ');
  const dst = ss.getSheetByName('発注');
  if (!src || !dst) { Logger.log('発注: シートが見つかりません'); return { added: 0, startRow: 0 }; }

  const srcVals = src.getDataRange().getValues();
  const dstVals = dst.getDataRange().getValues();

  // 見出し行（C列が DorderDate の行）
  let dstHeader = -1;
  for (let i = 0; i < dstVals.length; i++) {
    if (String(dstVals[i][2]).trim() === 'DorderDate') { dstHeader = i; break; }
  }
  if (dstHeader < 0) { Logger.log('発注: 見出し行が見つかりません'); return { added: 0, startRow: 0 }; }

  // 既存行のキー集合と最終データ行
  const byCode = {}, byName = {};
  let lastDataRow = dstHeader;
  for (let i = dstHeader + 1; i < dstVals.length; i++) {
    const code = String(dstVals[i][11]).trim(); // L列 商品コード
    const name = normText_(dstVals[i][9]);      // J列 商品名
    if (!code && !name) continue;
    lastDataRow = i;
    const o = normOrderNo_(dstVals[i][6]);      // G列 購入No.
    if (code) byCode[o + '|' + normCode_(code)] = true;
    if (name) byName[o + '|' + name] = true;
  }

  // 同期データ側の見出し（B列が DorderDate の行）
  let srcHeader = -1;
  for (let i = 0; i < srcVals.length; i++) {
    if (String(srcVals[i][1]).trim() === 'DorderDate') { srcHeader = i; break; }
  }
  if (srcHeader < 0) { Logger.log('同期データ: 見出し行が見つかりません'); return { added: 0, startRow: 0 }; }

  // 差分抽出
  const newRows = [];
  for (let i = srcHeader + 1; i < srcVals.length; i++) {
    const r = srcVals[i];
    const code = String(r[10]).trim();  // K列 商品コード
    const name = normText_(r[8]);       // I列 商品名
    if (!code && !name) continue;
    const o = normOrderNo_(r[5]);       // F列 発注NO
    if ((code && byCode[o + '|' + normCode_(code)]) ||
        (name && byName[o + '|' + name])) continue;
    newRows.push(r);
  }
  if (newRows.length === 0) return { added: 0, startRow: 0 };

  // 列マッピング [同期データの列(0始), 発注の列(1始)]
  // 関数列（A:チェック / B:No. / R:小計 / T:支払金額 / W〜AB）には書き込まない
  const MAP = [
    [1, 3],   // 発注日
    [2, 4],   // 入荷日
    [3, 5],   // 入荷数
    [4, 6],   // パッキング
    [5, 7],   // 購入No.
    [6, 8],   // 予約
    [7, 9],   // 業者
    [8, 10],  // 商品名
    [9, 11],  // オプション
    [10, 12], // 商品コード
    [11, 13], // 数量
    [12, 14], // 列1
    [13, 15], // 品目
    [14, 16], // ★重さ weight(g)（Excel O列 → 発注 P列）
    [15, 17], // 価格
    [19, 21], // 決済方法
    [20, 22], // 決済日
    [23, 29]  // 販売価格 (X→AC)
  ];

  const startRow = lastDataRow + 2; // 1始まりの追記開始行

  MAP.forEach(([s, d]) => {
    dst.getRange(startRow, d, newRows.length, 1)
       .setValues(newRows.map(r => [r[s]]));
  });

  Logger.log('発注に追加: ' + newRows.map(r => r[5] + '/' + r[10]).join(', '));
  return { added: newRows.length, startRow: startRow };
}

// ---- 見出し行から列を探す（0始まり、見つからなければ -1）----
function findColIndexByHeader_(headerRow, names) {
  const norm = v => String(v || '').trim().toLowerCase()
    .replace(/[\s\u3000]/g, '').replace(/[()（）]/g, '');
  const targets = names.map(norm);
  for (let c = 0; c < headerRow.length; c++) {
    if (targets.indexOf(norm(headerRow[c])) >= 0) return c;
  }
  return -1;
}

const WEIGHT_HEADERS_ = ['重さ', 'weight(g)', 'weight', '重量'];

// ---- EMS同期データ → EMSリスト（差分追加）----
function syncDiffToEmsList_(ss) {
  const src = ss.getSheetByName('EMS同期データ');
  const dst = ss.getSheetByName('EMSリスト');
  if (!src || !dst) { Logger.log('EMS: シートが見つかりません'); return { added: 0, startRow: 0 }; }

  const srcVals = src.getDataRange().getValues();
  const dstVals = dst.getDataRange().getValues();

  let dstHeader = -1;
  for (let i = 0; i < dstVals.length; i++) {
    if (String(dstVals[i][0]).trim() === 'No.') { dstHeader = i; break; }
  }
  if (dstHeader < 0) { Logger.log('EMSリスト: 見出し行が見つかりません'); return { added: 0, startRow: 0 }; }

  const byCode = {}, byQty = {}, byFallback = {}, byArrivalCodeQty = {};
  let lastDataRow = dstHeader, maxNo = 0, latestDstArrival = '';
  for (let i = dstHeader + 1; i < dstVals.length; i++) {
    const track = normTrack_(dstVals[i][12]);
    const code = normCode_(dstVals[i][8]);
    if (!track && !code) continue;
    lastDataRow = i;
    const dstArrival = _emsFmtDate_(dstVals[i][1]);   // B列 入荷日
    if (_emsIsDateKey_(dstArrival) && dstArrival > latestDstArrival) latestDstArrival = dstArrival;
    if (track && code) {
      codeKeys_(code).forEach(k => { byCode[track + '|' + k] = true; });
    } else if (code) {                                // ★EMS番号が空の迷子行
      const arr = dstArrival;
      const qty = String(dstVals[i][9]);              // J列 数量
      codeKeys_(code).forEach(k => { byFallback[arr + '|' + k + '|' + qty] = true; });
    }
    if (code) {
      const arr = dstArrival;
      const qty = String(dstVals[i][9]);               // J列 数量
      codeKeys_(code).forEach(k => { byArrivalCodeQty[arr + '|' + k + '|' + qty] = true; });
    }
    byQty[track + '|' + String(dstVals[i][9]) + '|' + normText_(dstVals[i][10]).toLowerCase()] = true;
    const n = Number(dstVals[i][0]);
    if (n > maxNo) maxNo = n;
  }

  let srcHeader = -1;
  for (let i = 0; i < srcVals.length; i++) {
    if (String(srcVals[i][0]).trim() === '入荷') { srcHeader = i; break; }
  }
  if (srcHeader < 0) { Logger.log('EMS同期データ: 見出し行が見つかりません'); return { added: 0, startRow: 0 }; }

  // ★重さ列を見出しから自動検出（どちらかに無ければ転記しない）
  const srcWeightIdx = findColIndexByHeader_(srcVals[srcHeader], WEIGHT_HEADERS_);          // 0始まり
  const dstWeightCol = findColIndexByHeader_(dstVals[dstHeader], WEIGHT_HEADERS_) + 1;      // 1始まり（0なら無し）

  const newRows = [];
  for (let i = srcHeader + 1; i < srcVals.length; i++) {
    const r = srcVals[i];
    const track = normTrack_(r[3]);
    const code = normCode_(r[7]);
    if (!track && !code) continue;
    const _arr = _emsFmtDate_(r[0]), _qty = String(r[8]);
    if (latestDstArrival && _emsIsDateKey_(_arr) && _arr < latestDstArrival) continue;
    // ★Excel側コードの別名のどれかが既存にあれば「既存」とみなす
    const hitCode = codeKeys_(code).some(k => byCode[track + '|' + k]);
    const hitFallback = codeKeys_(code).some(k => byFallback[_arr + '|' + k + '|' + _qty]);
    const hitBlankTrackExisting = !track && codeKeys_(code).some(k => byArrivalCodeQty[_arr + '|' + k + '|' + _qty]);
    const kQty = track + '|' + String(r[8]) + '|' + normText_(r[10]).toLowerCase();
    if (hitCode || hitFallback || hitBlankTrackExisting || byQty[kQty]) continue;
    newRows.push(r);
  }
  if (newRows.length === 0) return { added: 0, startRow: 0 };

  const startRow = lastDataRow + 2;

  dst.getRange(startRow, 1, newRows.length, 1)
     .setValues(newRows.map((_, k) => [maxNo + k + 1]));          // No.
  dst.getRange(startRow, 2, newRows.length, 1)
     .setValues(newRows.map(r => [r[0]]));                        // 入荷日
  dst.getRange(startRow, 3, newRows.length, 1)
     .setValues(newRows.map(r => [r[1]]));                        // EMS発送日
  dst.getRange(startRow, 4, newRows.length, 1)
     .setValues(newRows.map(() => ['⇒']));                        // ⇒
  dst.getRange(startRow, 5, newRows.length, 1)
     .setValues(newRows.map(r => [r[2]]));                        // EMS到着日
  dst.getRange(startRow, 9, newRows.length, 1)
     .setValues(newRows.map(r => [normCode_(r[7])]));             // 商品コード
  dst.getRange(startRow, 10, newRows.length, 1)
     .setValues(newRows.map(r => [r[8]]));                        // 数量
  // K列 品目はEMSリスト側の関数で出すので書き込まない
  dst.getRange(startRow, 13, newRows.length, 1)
     .setValues(newRows.map(r => [normTrack_(r[3])]));            // EMS番号

  // ★重さ（両方のシートに重さ列があるときだけ）
  if (srcWeightIdx >= 0 && dstWeightCol >= 1) {
    dst.getRange(startRow, dstWeightCol, newRows.length, 1)
       .setValues(newRows.map(r => [r[srcWeightIdx]]));
  }

  Logger.log('EMSリストに追加: ' + newRows.map(r => normTrack_(r[3])).join(', '));
  return { added: newRows.length, startRow: startRow };
}

// ---- EMS同期データ → EMSリスト既存行の更新（入荷日・発送日・到着日）----
function updateEmsListFromSync_(ss) {
  const src = ss.getSheetByName('EMS同期データ');
  const dst = ss.getSheetByName('EMSリスト');
  if (!src || !dst) return 0;

  const srcVals = src.getDataRange().getValues();
  const dstVals = dst.getDataRange().getValues();

  let dstHeader = -1;
  for (let i = 0; i < dstVals.length; i++) {
    if (String(dstVals[i][0]).trim() === 'No.') { dstHeader = i; break; }
  }
  if (dstHeader < 0) return 0;

  // ① EMS番号＋商品コード → 行（ひも付け済み）
  const rowByKey = {};
  // ② 入荷日＋コード＋数量 → 行（EMS番号が空の迷子行だけ）
  const rowByFallback = {};
  // ③ コード＋数量 → 行（日付やEMS番号まで空で入ってしまった行の補修用）
  const rowByLoose = {};
  for (let i = dstHeader + 1; i < dstVals.length; i++) {
    const track = normTrack_(dstVals[i][12]);  // M列 EMS番号
    const code  = normCode_(dstVals[i][8]);    // I列 商品コード
    if (!code) continue;
    const qty = String(dstVals[i][9]);          // J列 数量
    if (track) {
      codeKeys_(code).forEach(k => {
        const key = track + '|' + k;
        if (!(key in rowByKey)) rowByKey[key] = i;
      });
    } else {
      const arr = _emsFmtDate_(dstVals[i][1]); // B列 入荷日
      codeKeys_(code).forEach(k => {
        const key = arr + '|' + k + '|' + qty;
        if (!(key in rowByFallback)) rowByFallback[key] = i;
        const looseKey = k + '|' + qty;
        (rowByLoose[looseKey] = rowByLoose[looseKey] || []).push(i);
      });
    }
  }

  let srcHeader = -1;
  for (let i = 0; i < srcVals.length; i++) {
    if (String(srcVals[i][0]).trim() === '入荷') { srcHeader = i; break; }
  }
  if (srcHeader < 0) return 0;

  const UPDATE_COLS = [
    [0, 2, '入荷日'],   // A → B
    [1, 3, '発送日'],   // B → C
    [2, 5, '到着日']    // C → E
  ];

  const usedFallback = {};
  let updates = 0;

  for (let i = srcHeader + 1; i < srcVals.length; i++) {
    const r = srcVals[i];
    const track = normTrack_(r[3]);   // D列 Tracking #
    const code  = normCode_(r[7]);    // H列 ItemCode
    if (!code) continue;

    // まずEMS番号でひも付け済みの行
    let row;
    if (track) {
      for (const k of codeKeys_(code)) {
        if (rowByKey[track + '|' + k] !== undefined) { row = rowByKey[track + '|' + k]; break; }
      }
    }

    // 無かったら、EMS番号が空の迷子行を 入荷日＋コード＋数量 で探す
    let isFallback = false;
    if (row === undefined) {
      const arr = _emsFmtDate_(r[0]);  // A列 入荷
      const qty = String(r[8]);        // I列 Qty
      for (const k of codeKeys_(code)) {
        const key = arr + '|' + k + '|' + qty;
        if (rowByFallback[key] !== undefined && !usedFallback[key]) {
          row = rowByFallback[key]; isFallback = true; usedFallback[key] = true; break;
        }
      }
    }

    // 日付やEMS番号まで空で入ってしまった既存行は、コード＋数量で補修する
    let isLoose = false;
    if (row === undefined) {
      const qty = String(r[8]);        // I列 Qty
      for (const k of codeKeys_(code)) {
        const key = k + '|' + qty;
        const rows = rowByLoose[key];
        if (rows && rows.length > 0) {
          row = rows.shift(); isLoose = true; break;
        }
      }
    }
    if (row === undefined) continue;

    UPDATE_COLS.forEach(([s, d]) => {
      const v = r[s];
      if (!isBlank_(v) && !sameVal_(v, dstVals[row][d - 1])) {
        dst.getRange(row + 1, d).setValue(v);
        dstVals[row][d - 1] = v;
        updates++;
      }
    });

    // 迷子行や空欄補修行やったらEMS番号(M=13)も埋める
    if ((isFallback || isLoose) && track && isBlank_(dstVals[row][12])) {
      dst.getRange(row + 1, 13).setValue(track);
      dstVals[row][12] = track;
      updates++;
    }
  }
  return updates;
}

function EMSリスト_空欄行をEMS同期データから補完(silent) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('EMS同期データ');
  const dst = ss.getSheetByName('EMSリスト');
  if (!src || !dst) return 0;

  const srcVals = src.getDataRange().getValues();
  const dstVals = dst.getDataRange().getValues();

  let srcHeader = -1;
  for (let i = 0; i < srcVals.length; i++) {
    if (String(srcVals[i][0]).trim() === '入荷') { srcHeader = i; break; }
  }
  let dstHeader = -1;
  for (let i = 0; i < dstVals.length; i++) {
    if (String(dstVals[i][0]).trim() === 'No.') { dstHeader = i; break; }
  }
  if (srcHeader < 0 || dstHeader < 0) return 0;

  const srcByLoose = {};
  for (let i = srcHeader + 1; i < srcVals.length; i++) {
    const r = srcVals[i];
    const code = normCode_(r[7]);     // H列 ItemCode
    const qty = String(r[8]);         // I列 Qty
    if (!code || !qty) continue;
    if (isBlank_(r[0]) && isBlank_(r[1]) && isBlank_(r[2]) && !normTrack_(r[3])) continue;
    codeKeys_(code).forEach(k => {
      const key = k + '|' + qty;
      (srcByLoose[key] = srcByLoose[key] || []).push(r);
    });
  }

  let updates = 0;
  const touchedRows = [];
  for (let i = dstHeader + 1; i < dstVals.length; i++) {
    const row = dstVals[i];
    const code = normCode_(row[8]);   // I列 商品コード
    const qty = String(row[9]);       // J列 数量
    if (!code || !qty) continue;

    const needsFill =
      isBlank_(row[1]) ||             // B列 入荷日
      isBlank_(row[2]) ||             // C列 EMS発送日
      isBlank_(row[12]);              // M列 EMS番号
    if (!needsFill) continue;

    let srcRow;
    for (const k of codeKeys_(code)) {
      const rows = srcByLoose[k + '|' + qty];
      if (rows && rows.length > 0) { srcRow = rows.shift(); break; }
    }
    if (!srcRow) continue;

    const writes = [
      [1, 2, srcRow[0]],              // A 入荷 -> B
      [2, 3, srcRow[1]],              // B 出発 -> C
      [4, 5, srcRow[2]],              // C 到着 -> E
      [12, 13, normTrack_(srcRow[3])] // D Tracking # -> M
    ];

    let rowTouched = false;
    writes.forEach(([idx, col, val]) => {
      if (!isBlank_(val) && isBlank_(dstVals[i][idx])) {
        dst.getRange(i + 1, col).setValue(val);
        dstVals[i][idx] = val;
        updates++;
        rowTouched = true;
      }
    });

    if (rowTouched) touchedRows.push(i + 1);
  }

  if (touchedRows.length && typeof EMS_updateDatesByRows_ === 'function') {
    const minRow = Math.min(...touchedRows);
    const maxRow = Math.max(...touchedRows);
    EMS_updateDatesByRows_(dst, minRow, maxRow - minRow + 1, false);
  }

  const msg = `EMS空欄補完: ${updates}セル`;
  if (silent) Logger.log(msg);
  else SpreadsheetApp.getActive().toast(msg);
  return updates;
}
// ---- 既存行の入荷日・入荷数・重さを同期データ側の値で更新 ----
function updateArrivalsInHatchu_(ss) {
  const src = ss.getSheetByName('発注リスト同期データ');
  const dst = ss.getSheetByName('発注');
  if (!src || !dst) return 0;

  const srcVals = src.getDataRange().getValues();
  const dstVals = dst.getDataRange().getValues();

  // 発注側の見出し行
  let dstHeader = -1;
  for (let i = 0; i < dstVals.length; i++) {
    if (String(dstVals[i][2]).trim() === 'DorderDate') { dstHeader = i; break; }
  }
  if (dstHeader < 0) return 0;

  // 発注側のキー → 行番号（0始まり）の対応表
  const rowByCode = {}, rowByName = {};
  for (let i = dstHeader + 1; i < dstVals.length; i++) {
    const code = String(dstVals[i][11]).trim(); // L列 商品コード
    const name = normText_(dstVals[i][9]);      // J列 商品名
    if (!code && !name) continue;
    const o = normOrderNo_(dstVals[i][6]);      // G列 購入No.
    if (code && !(o + '|' + normCode_(code) in rowByCode)) rowByCode[o + '|' + normCode_(code)] = i;
    if (name && !(o + '|' + name in rowByName)) rowByName[o + '|' + name] = i;
  }

  // 同期データ側の見出し行
  let srcHeader = -1;
  for (let i = 0; i < srcVals.length; i++) {
    if (String(srcVals[i][1]).trim() === 'DorderDate') { srcHeader = i; break; }
  }
  if (srcHeader < 0) return 0;

  // Excel側に値があって発注側と違うときだけ上書きする項目
  // [同期データの列(0始), 発注の列(1始), 名前]
  const UPDATE_COLS = [
    [2, 4, '入荷日'],    // C列 → D列
    [3, 5, '入荷数'],    // D列 → E列
    [14, 16, '重さ']     // O列 → P列 ★追加
  ];

  let updates = 0;
  for (let i = srcHeader + 1; i < srcVals.length; i++) {
    const r = srcVals[i];

    // 更新対象の値がひとつも無い行はスキップ
    if (UPDATE_COLS.every(([s]) => isBlank_(r[s]))) continue;

    const o = normOrderNo_(r[5]);          // F列 発注NO
    const code = normCode_(r[10]);         // K列 商品コード
    const name = normText_(r[8]);          // I列 商品名
    let row = rowByCode[o + '|' + code];
    if (row === undefined) row = rowByName[o + '|' + name];
    if (row === undefined) continue;       // 発注側に行がない（追加処理側で入る）

    UPDATE_COLS.forEach(([s, d]) => {
      const v = r[s];
      if (!isBlank_(v) && !sameVal_(v, dstVals[row][d - 1])) {
        dst.getRange(row + 1, d).setValue(v);
        updates++;
      }
    });
  }
  return updates;
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '' || String(v).trim() === '　';
}

function sameVal_(a, b) {
  const f = v => {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (v === null || v === undefined) return '';
    return String(v).trim();
  };
  return f(a) === f(b);
}

// ---- 発注シート：onEdit相当の後処理 ----
function runHatchuPostProcess_(sh, info) {
  // 追加行まわりの罫線とチェックボックス
  if (info.added > 0 &&
      typeof HATCHU_processBordersAndCheckboxes_ === 'function' &&
      typeof HATCHU_CFG !== 'undefined') {
    HATCHU_processBordersAndCheckboxes_(sh, info.startRow, info.added);
  }

  // 購入Noごとのグループ太枠
  if (typeof 発注_drawGroupBorders_ === 'function') {
    発注_drawGroupBorders_(sh);
  }

  // 消込の色付け
  SpreadsheetApp.flush();
  if (typeof colorKeshikomiAllRows_ === 'function') {
    colorKeshikomiAllRows_(sh);
  }

  // 商品マスタへの自動登録（未登録のコード＋業者を一括追記、silent実行）
  if (typeof マスタ自動補完_ === 'function') {
    マスタ自動補完_(true);
  }
}

// ---- EMSリスト：onEdit相当の後処理 ----
function runEmsPostProcess_(sh, info) {
  if (typeof EMS_CFG === 'undefined' || info.added <= 0) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return;
  try {
    // 発送日列を編集したのと同じ扱いでステータスを更新
    if (typeof EMS_updateStatusOnlyForEditedRows_ === 'function') {
      const r = sh.getRange(info.startRow, EMS_CFG.COL_EMS_DATE, info.added, 1);
      EMS_updateStatusOnlyForEditedRows_(sh, r);
    }
    // 日付・Boxまわりの更新（onEditのAUTO_BOX処理と同じ）
    if (typeof EMS_updateDatesByRows_ === 'function') {
      EMS_updateDatesByRows_(sh, info.startRow, info.added, false);
    }
  } finally {
    lock.releaseLock();
  }
}

// ====== EMSリスト → Googleカレンダー同期 ======
const EMS_CAL_CFG = {
  CALENDAR_NAME: 'EMS追跡',
  MARKER: '[EMS-SYNC]',  // スクリプトが作ったイベントの目印
  COLORS: {
    '未着':       CalendarApp.EventColor.RED,
    '輸送中':     CalendarApp.EventColor.ORANGE,
    '在庫反映済み': CalendarApp.EventColor.GREEN
  }
};

function syncEmsCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('EMSリスト');
  if (!sh) return;

  const vals = sh.getDataRange().getValues();
  let header = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') { header = i; break; }
  }
  if (header < 0) return;

  // EMS番号＋Box単位でグループ化
  const groups = {};
  for (let i = header + 1; i < vals.length; i++) {
    const r = vals[i];
    const track = normTrack_(r[12]);              // M列 EMS番号
    if (!track) continue;
    const ship = r[2], arrive = r[4];             // C列 発送日 / E列 到着日
    if (!(ship instanceof Date)) continue;        // 発送日が無い行はスキップ
    const key = track + '|' + String(r[13]);      // Box No.込みでキー化
    if (!groups[key]) {
      groups[key] = {
        track: track,
        box: r[13],
        ship: ship,
        arrive: (arrive instanceof Date) ? arrive : null,
        status: String(r[6] || '').trim(),        // G列 ステータス
        items: [], qty: 0
      };
    }
    groups[key].items.push(`${r[8]} ×${r[9]}（${r[10]}）`); // 商品コード×数量（品目）
    groups[key].qty += Number(r[9]) || 0;
    // 1行でも「未着」があれば箱全体を未着扱い
    if (String(r[6]).trim() === '未着') groups[key].status = '未着';
  }

  // 専用カレンダーを取得（なければ作る）
  let cal = CalendarApp.getCalendarsByName(EMS_CAL_CFG.CALENDAR_NAME)[0];
  if (!cal) cal = CalendarApp.createCalendar(EMS_CAL_CFG.CALENDAR_NAME);

  // スクリプト製の既存イベントを全削除して作り直す（常にシートが正）
  cal.getEvents(new Date('2026-01-01'), new Date('2031-01-01')).forEach(ev => {
    if (ev.getDescription().indexOf(EMS_CAL_CFG.MARKER) >= 0) ev.deleteEvent();
  });

  // 箱ごとに「発送日〜到着日」の帯イベントを作成
  // 箱ごとに「発送日〜到着日」の帯イベントを作成
  let count = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);   // ★追加
  Object.keys(groups).forEach(key => {
    const g = groups[key];
    const start = g.ship;
    const endBase = g.arrive || g.ship;
    const end = new Date(endBase.getFullYear(), endBase.getMonth(), endBase.getDate() + 1);

    const title = `【${g.status || '状態不明'}】${g.track} Box${g.box}（${g.items.length}品目 ${g.qty}点）`;
    const desc = EMS_CAL_CFG.MARKER + '\n' +
      '追跡: https://trackings.post.japanpost.jp/services/srv/search/input\n' +
      g.items.join('\n');

    const ev = cal.createAllDayEvent(title, start, end, { description: desc });
    const color = (endBase < today)                        // ★到着済みはグレー
      ? CalendarApp.EventColor.GRAY
      : EMS_CAL_CFG.COLORS[g.status];
    if (color) ev.setColor(color);
    count++;
  });

  Logger.log(`EMSカレンダー同期完了: ${count}箱`);
}

// ---- EMSリストを箱単位にまとめる（カレンダー系の共通処理）----
function getEmsBoxGroups_(ss) {
  const sh = ss.getSheetByName('EMSリスト');
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  let header = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') { header = i; break; }
  }
  if (header < 0) return {};

  // M列（EMS番号）の背景色＝EMSリストの箱グループ色を帯にも使う
  const emsBgs = sh.getRange(1, 13, vals.length, 1).getBackgrounds();

  const groups = {};
  for (let i = header + 1; i < vals.length; i++) {
    const r = vals[i];
    const track = normTrack_(r[12]);
    if (!track) continue;
    const ship = r[2], arrive = r[4];
    if (!(ship instanceof Date)) continue;
    const key = track + '|' + String(r[13]);
    if (!groups[key]) {
      groups[key] = {
        track: track, box: r[13], ship: ship,
        arrive: (arrive instanceof Date) ? arrive : null,
        status: String(r[6] || '').trim(),
        color: emsBgs[i][0],
        items: [], qty: 0
      };
    }
    groups[key].items.push(`${r[8]} ×${r[9]}（${r[10]}）`);
    groups[key].qty += Number(r[9]) || 0;
    if (String(r[6]).trim() === '未着') groups[key].status = '未着';
  }
  return groups;
}

// 表示する期間（今日を基準に前後何ヶ月か）
const EMS_CAL_VIEW = { MONTHS_BACK: 1, MONTHS_FWD: 1 };

function buildEmsCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groups = getEmsBoxGroups_(ss);
  let boxes = Object.keys(groups).map(k => groups[k]);
  if (boxes.length === 0) { ss.toast('EMSデータがありません'); return; }

  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const STYLE = {
    '未着':         { bg: '#d93025', font: '#ffffff' },
    '輸送中':       { bg: '#f29900', font: '#ffffff' },
    '在庫反映済み': { bg: '#b7dec7', font: '#3c4043' },
    'default':      { bg: '#dadce0', font: '#3c4043' }
  };
  const PAST_STYLE = { bg: '#e8eaed', font: '#9aa0a6' }; // 到着済みのグレーアウト

  const today = new Date(); today.setHours(0, 0, 0, 0);

  // 表示ウィンドウ（先月初〜2ヶ月先の月末）
  const winStart = new Date(today.getFullYear(), today.getMonth() - EMS_CAL_VIEW.MONTHS_BACK, 1);
  const winEnd = new Date(today.getFullYear(), today.getMonth() + EMS_CAL_VIEW.MONTHS_FWD + 1, 0);

  // 箱の整形＋ウィンドウ外を除外
  boxes.forEach(b => { b.end = b.arrive || b.ship; });
  boxes = boxes.filter(b => b.end >= winStart && b.ship <= winEnd);

  boxes.forEach(b => {
    b.isPast = b.end < today;
    if (b.isPast) {
      // 到着済みは今まで通りグレー
      b.style = PAST_STYLE;
    } else {
      // EMSリストのM列の色をそのまま帯に使う（色が無ければステータス色にフォールバック）
      const hasColor = b.color && String(b.color).toLowerCase() !== '#ffffff';
      b.style = hasColor
        ? { bg: b.color, font: '#3c4043' }
        : (STYLE[b.status] || STYLE['default']);
    }
    b.note = `【${b.status || '?'}】${b.track} Box${b.box}\n` +
             `発送: ${fmt(b.ship)} → 到着: ${fmt(b.end)}\n` +
             `${b.qty}点\n` + b.items.join('\n');
  });

  // レーン割当（重なる箱は別の段へ）
  boxes.sort((a, b) => a.ship - b.ship);
  const laneEnd = [];
  boxes.forEach(b => {
    let lane = laneEnd.findIndex(e => e < b.ship.getTime());
    if (lane < 0) { lane = laneEnd.length; laneEnd.push(0); }
    laneEnd[lane] = b.end.getTime();
    b.lane = lane;
  });

  // 帯の長さに応じたラベル（ステータスは常に先頭に表示）
  const bandLabel = (b, days) => {
    if (days >= 4) return `【${b.status || '?'}】${b.track} B${b.box}（${b.qty}点）`;
    if (days >= 2) return `【${b.status || '?'}】${b.track} B${b.box}`;
    return `【${b.status || '?'}】${b.track}`;
  };

  let sh = ss.getSheetByName('EMSカレンダー');
  if (!sh) sh = ss.insertSheet('EMSカレンダー');
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearNotes();

  const youbi = ['日', '月', '火', '水', '木', '金', '土'];
  let row = 1;
  let cur = new Date(winStart);

  while (cur <= winEnd) {
    const y = cur.getFullYear(), mo = cur.getMonth();
    const monthStart = new Date(y, mo, 1);
    const daysInMonth = new Date(y, mo + 1, 0).getDate();
    const monthEnd = new Date(y, mo, daysInMonth);
    const blockStart = row;

    sh.getRange(row, 1, 1, 7).merge().setValue(`${y}年${mo + 1}月`)
      .setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#f1f3f4');
    row++;

    sh.getRange(row, 1, 1, 7).setValues([youbi])
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setBackground('#fafafa');
    sh.getRange(row, 1).setFontColor('#d93025');
    sh.getRange(row, 7).setFontColor('#1a73e8');
    row++;

    let weekStart = new Date(y, mo, 1);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay());

    while (weekStart <= monthEnd) {
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);

      // 日付番号行（他の月の日はグレー薄表示）
      const numRow = row;
      for (let c = 0; c < 7; c++) {
        const d = new Date(weekStart); d.setDate(d.getDate() + c);
        const cell = sh.getRange(numRow, c + 1);
        if (d.getMonth() === mo) {
          cell.setValue(d.getDate());
          if (fmt(d) === fmt(today)) {
            cell.setBackground('#fde293').setFontWeight('bold').setFontColor('#3c4043');
          }
        } else {
          cell.setBackground('#f8f9fa');
        }
      }
      sh.getRange(numRow, 1, 1, 7)
        .setFontSize(14)
        .setHorizontalAlignment('right').setFontColor('#70757a');
      sh.setRowHeight(numRow, 26);
      row++;

      // この週×この月にかかる箱だけ描く（月またぎの二重描画を防止）
      const segs = [];
      boxes.forEach(b => {
        let s = b.ship > weekStart ? b.ship : weekStart;
        let e = b.end < weekEnd ? b.end : weekEnd;
        if (s < monthStart) s = monthStart;
        if (e > monthEnd) e = monthEnd;
        if (s > e) return;
        segs.push({ b: b, s: s, e: e });
      });

      const maxLane = segs.reduce((m, x) => Math.max(m, x.b.lane), -1);
      const laneRows = Math.max(maxLane + 1, 1);
    for (let L = 0; L < laneRows; L++) sh.setRowHeight(row + L, 32);  // ★32に

      segs.forEach(x => {
        const c1 = Math.floor((x.s - weekStart) / 86400000) + 1;
        const c2 = Math.floor((x.e - weekStart) / 86400000) + 1;
        const days = c2 - c1 + 1;
        const range = sh.getRange(row + x.b.lane, c1, 1, days);
        if (days > 1) range.merge();
        range.setValue(bandLabel(x.b, days))
          .setBackground(x.b.style.bg).setFontColor(x.b.style.font)
          .setFontSize(12)   // ★12ptに
          .setVerticalAlignment('middle')
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
          .setNote(x.b.note);
      });

      row += laneRows;
      sh.setRowHeight(row, 8);
      row++;

      weekStart = new Date(weekEnd);
      weekStart.setDate(weekStart.getDate() + 1);
    }

    sh.getRange(blockStart, 1, row - blockStart, 7)
      .setBorder(true, true, true, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    row++;
    cur = new Date(y, mo + 1, 1);
  }

 for (let c = 1; c <= 7; c++) sh.setColumnWidth(c, 180);  // ★180に
  sh.setHiddenGridlines(true);
  ss.toast('EMSカレンダーを更新したで');
}

function Googleカレンダーをシートに同期() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName('Googleカレンダー同期');
  if (!sh) {
    sh = ss.insertSheet('Googleカレンダー同期');
  }

  const calendarName = 'EMS追跡';
  const calendars = CalendarApp.getCalendarsByName(calendarName);

  if (calendars.length === 0) {
    SpreadsheetApp.getUi().alert(`Googleカレンダー「${calendarName}」が見つかりません。`);
    return;
  }

  const cal = calendars[0];

  const startDate = new Date();
  startDate.setMonth(startDate.getMonth() - 1);

  const endDate = new Date();
  endDate.setMonth(endDate.getMonth() + 3);

  const events = cal.getEvents(startDate, endDate);

  const values = [
    ['開始日', '終了日', 'タイトル', '説明', '場所', 'カレンダー名']
  ];

  events.forEach(ev => {
    values.push([
      ev.getStartTime(),
      ev.getEndTime(),
      ev.getTitle(),
      ev.getDescription(),
      ev.getLocation(),
      calendarName
    ]);
  });

  sh.clearContents();

  sh.getRange(1, 1, values.length, values[0].length)
    .setValues(values);

  sh.getRange(1, 1, 1, values[0].length)
    .setBackground('#ffff00')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (values.length > 1) {
    sh.getRange(2, 1, values.length - 1, 2)
      .setNumberFormat('yyyy/m/d');
  }

  sh.autoResizeColumns(1, values[0].length);

  SpreadsheetApp.getActive().toast(
    `Googleカレンダーをシートに同期しました：${events.length}件`
  );
}

function EMS同期データ_3行目基準で列幅調整() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('EMS同期データ');

  if (!sh) {
    SpreadsheetApp.getUi().alert('EMS同期データが見つからんで');
    return;
  }

  const colCount = sh.getLastColumn();

  sh.getDataRange()
    .setFontSize(15)
    .setVerticalAlignment('middle')
    .setWrap(false);

  sh.setFrozenRows(2);
  fitColumnsBySampleRow_(sh, 3, colCount);

  ss.toast('EMS同期データを3行目基準で列幅調整したで');
}

// 商品コードの別名展開（KRSJCM03-0506-06 → KRSJCM03-06 も同一とみなす）
function codeKeys_(code) {
  const c = normCode_(code);
  const keys = [c];
  const m = c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if (m && (m[4] === m[2] || m[4] === m[3])) {
    keys.push(m[1] + '-' + m[4]);   // 短縮形を別名として追加
  }
  return keys;
}

// 日付を yyyy-MM-dd に正規化（"26/06/16(火)" みたいな文字列にも対応）
function _emsFmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (typeof v === 'number' && isFinite(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 24 * 60 * 60 * 1000);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  const s = String(v || '').trim().replace(/\(.+?\)/, '');
  const m = s.match(/^(\d{2,4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) { let y = Number(m[1]); if (y < 100) y += 2000;
    return y + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[3]).slice(-2); }
  const jp = s.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (jp) {
    const y = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
    return y + '-' + ('0'+jp[1]).slice(-2) + '-' + ('0'+jp[2]).slice(-2);
  }
  return s;
}

function _emsIsDateKey_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
}

function EMSリスト_重複行を確認して削除() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('EMSリスト');
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('EMSリストが無いで'); return; }

  const vals = sh.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('見出し行(No.)が見つからんで'); return; }

  // 入荷日＋商品コード＋数量 でグループ化
  const groups = {};
  for (let i = h + 1; i < vals.length; i++) {
    const code = normCode_(vals[i][8]);     // I列 商品コード
    if (!code) continue;
    const arr = _emsFmtDate_(vals[i][1]);   // B列 入荷日
    const qty = String(vals[i][9]);         // J列 数量
    const key = arr + '|' + code + '|' + qty;
    (groups[key] = groups[key] || []).push(i);
  }

  // 統合先の空欄に流し込む列：B C E F G H L M N
  const MERGE_COLS = [2, 3, 5, 6, 7, 8, 12, 13, 14];

  const plans = [];
  Object.keys(groups).forEach(key => {
    const rows = groups[key];
    if (rows.length < 2) return;
    const keep = Math.min(...rows);
    plans.push({ keep: keep, dels: rows.filter(r => r !== keep) });
  });

  if (plans.length === 0) { ui.alert('重複行は見つからんかったで'); return; }

  let preview = '', delCount = 0;
  plans.forEach(p => {
    delCount += p.dels.length;
    preview += `${normCode_(vals[p.keep][8])}: 行${p.keep+1}に統合 ← 行${p.dels.map(d=>d+1).join(',')}削除\n`;
  });

  const res = ui.alert(
    '重複行クリーンアップ',
    `重複グループ ${plans.length}件 / 削除 ${delCount}行\n\n` +
    preview.slice(0, 1200) +
    '\n統合先の空欄だけ、削除する行の値で埋めてから消すで。実行する？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { ui.alert('やめといたで'); return; }

  // 統合（keepの空欄だけ埋める）
  plans.forEach(p => {
    p.dels.forEach(d => {
      MERGE_COLS.forEach(col => {
        if (isBlank_(vals[p.keep][col-1]) && !isBlank_(vals[d][col-1])) {
          sh.getRange(p.keep + 1, col).setValue(vals[d][col-1]);
          vals[p.keep][col-1] = vals[d][col-1];
        }
      });
    });
  });

  // 行削除（下から）
  const delRows = [];
  plans.forEach(p => p.dels.forEach(d => delRows.push(d)));
  [...new Set(delRows)].sort((a,b)=>b-a).forEach(r => sh.deleteRow(r + 1));

  ui.alert(`完了: ${delRows.length}行 削除したで`);
}

function EMSリスト_下に混ざった古い行を確認して削除() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('EMSリスト');
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('EMSリストが無いで'); return; }

  const vals = sh.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('見出し行(No.)が見つからんで'); return; }

  let latestSeen = '';
  const targets = [];
  for (let i = h + 1; i < vals.length; i++) {
    const arr = _emsFmtDate_(vals[i][1]); // B列 入荷日
    if (!_emsIsDateKey_(arr)) continue;
    if (latestSeen && arr < latestSeen) {
      targets.push({
        row: i + 1,
        arr: arr,
        purchase: String(vals[i][5] || ''),
        code: normCode_(vals[i][8]),
        qty: String(vals[i][9] || ''),
        track: normTrack_(vals[i][12])
      });
      continue;
    }
    if (arr > latestSeen) latestSeen = arr;
  }

  if (targets.length === 0) {
    ui.alert('下に混ざった古い行は見つからんかったで');
    return;
  }

  const preview = targets
    .slice(0, 40)
    .map(t => `行${t.row}: ${t.arr} ${t.purchase} ${t.code} x${t.qty} ${t.track}`)
    .join('\n');
  const more = targets.length > 40 ? `\n...ほか ${targets.length - 40}行` : '';
  const res = ui.alert(
    '下に混ざった古いEMS行を削除',
    `新しい入荷日の下にある古い行を ${targets.length}行 見つけました。\n\n${preview}${more}\n\n削除する？`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { ui.alert('やめといたで'); return; }

  targets.map(t => t.row).sort((a,b)=>b-a).forEach(row => sh.deleteRow(row));
  ui.alert(`完了: ${targets.length}行 削除したで`);
}
