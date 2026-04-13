/**
 * アラジン.gs
 * アラジン OpenAPI を使った商品情報自動取得
 *
 * 【設定方法】
 *   スクリプトプロパティ（PropertiesService）に以下を登録：
 *     ALADIN_TTB_KEY = あなたのTTBキー
 *
 * 【発火タイミング】
 *   各シートで「購入URL」列にアラジンURLを入力
 *   または「ISBN」列にISBN13を入力
 *   → onEditで自動検出 → API呼び出し → 各列に自動入力
 */

// ============================================================
// TTBキー管理
// ============================================================
function アラジンTTBキーを取得_() {
  return PropertiesService.getScriptProperties().getProperty('ALADIN_TTB_KEY') || '';
}

function アラジンTTBキーを設定() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('アラジン TTBキー設定', 'TTBキーを入力してください:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) { ui.alert('キーが空です'); return; }
  PropertiesService.getScriptProperties().setProperty('ALADIN_TTB_KEY', key);
  ui.alert('✅ TTBキーを保存しました');
}

// ============================================================
// アラジンAPI呼び出し（ItemId または ISBN13）
// ============================================================
function アラジンAPI呼び出し_(idType, id) {
  const ttbKey = アラジンTTBキーを取得_();
  if (!ttbKey) throw new Error('TTBキーが設定されていません（メニュー → アラジンTTBキー設定）');

  const url = `https://www.aladin.co.kr/ttb/api/ItemLookUp.aspx`
    + `?ttbkey=${encodeURIComponent(ttbKey)}`
    + `&itemIdType=${idType}`
    + `&ItemId=${encodeURIComponent(id)}`
    + `&output=js`
    + `&Version=20131101`
    + `&Cover=Big`
    + `&OptResult=authors,fulldescription,subInfo`;

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error(`HTTP ${code}`);

    // アラジンAPIはJSON-Pっぽい形式で返ってくる場合があるため整形
    let text = resp.getContentText('utf-8');
    // "var _result = { ... };" 形式の場合は除去
    text = text.replace(/^\s*var\s+\w+\s*=\s*/, '').replace(/;\s*$/, '');
    const data = JSON.parse(text);
    if (!data.item || data.item.length === 0) return null;
    return data.item[0];
  } catch (e) {
    Logger.log('アラジンAPI エラー: ' + e.message);
    throw e;
  }
}

// URLからItemIdを抽出
function アラジンURLからItemIdを抽出_(url) {
  if (!url) return null;
  const m = String(url).match(/[?&]ItemId=(\d+)/i);
  return m ? m[1] : null;
}

// ============================================================
// ジャンル判定
// ============================================================
const ALADIN_CATEGORY_MAP = {
  // 音楽・映像
  'CD': '音楽映像', '음반': '音楽映像', 'Music': '音楽映像',
  'DVD': '音楽映像', 'Blu-ray': '音楽映像', '블루레이': '音楽映像',
  'LP': '音楽映像',
  // マンガ
  '만화': 'マンガ', 'Comic': 'マンガ', '코믹': 'マンガ',
  // 書籍
  '소설': '書籍', '에세이': '書籍', '시나리오': '書籍', '악보': '書籍',
  '잡지': '書籍', '문학': '書籍', '역사': '書籍',
  // グッズ
  'Gift': 'グッズ', 'Goods': 'グッズ', '굿즈': 'グッズ',
};

function アラジンジャンルを判定_(item) {
  if (!item) return 'その他';
  const cn = String(item.categoryName || '').toLowerCase();
  const ml = String(item.mallType || '').toLowerCase();
  const combined = cn + ' ' + ml;

  // mallTypeで大分類判定（最も確実）
  if (/music|음반|cd|lp/i.test(combined)) return '音楽映像';
  if (/dvd|bluray|blu-ray|블루레이|video/i.test(combined)) return '音楽映像';
  if (/comic|만화/i.test(combined)) return 'マンガ';
  if (/gift|goods|굿즈/i.test(combined)) return 'グッズ';
  if (/book|도서|novel|소설|essay|에세이|잡지|magazine/i.test(combined)) return '書籍';

  // categoryNameからフォールバック
  for (const [keyword, genre] of Object.entries(ALADIN_CATEGORY_MAP)) {
    if (combined.includes(keyword.toLowerCase())) return genre;
  }
  return '書籍'; // デフォルト
}

// ============================================================
// シート別列マッピング
// ============================================================
const ALADIN_COLUMN_MAP = {
  '韓国書籍': {
    trigger:          ['ISBN', 'アラジンURL'],
    url:              'アラジンURL',
    isbn:             'ISBN',
    title:            '商品名(原題)',
    author:           '著者',
    publisher:        '出版社',
    pubDate:          '発売日',
    price:            '原価',
    cover:            'メイン画像URL',
    additionalImages: '追加画像URL',
    description:      '商品説明',
    // categoryName 削除（手動プルダウンで選択）
    itemId:           'アラジン商品ID',
  },
  '韓国マンガ': {
    trigger:          ['ISBN', 'アラジンURL'],
    url:              'アラジンURL',
    isbn:             'ISBN',
    title:            '商品名(原題)',
    author:           '著者',
    publisher:        '出版社',
    pubDate:          '発売日',
    price:            '原価',
    cover:            'メイン画像URL',
    additionalImages: '追加画像URL',
    description:      '商品説明',
    // categoryName 削除（手動プルダウンで選択）
    itemId:           'アラジン商品ID',
  },
  '韓国音楽映像': {
  trigger:          ['アラジンURL'],
  url:              'アラジンURL',
  isbn:             'JANコード',
  title:            '商品名(原題)',
  author:           'アーティスト名',
  pubDate:          '発売日',
  price:            '原価',
  cover:            'メイン画像URL',
  additionalImages: '追加画像URL',
  description:      '商品説明',
  // categoryName は削除（手動プルダウンで選択）
  itemId:           'アラジン商品ID',
},
  '韓国グッズ': {
    trigger:          ['購入URL'],
    url:              '購入URL',
    title:            '商品名（原題）',
    pubDate:          '発売日',
    price:            '原価',
    cover:            'メイン画像URL',
    additionalImages: '追加画像URL',
    description:      '商品説明',
    itemId:           'アラジン商品ID',
  },
};

// ============================================================
// onEditから呼び出されるメイン処理
// ============================================================
function アラジン_onEdit(e) {
  const sh = e.range.getSheet();
  const shName = sh.getName();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const colMap = ALADIN_COLUMN_MAP[shName];
  if (!colMap) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  const col = e.range.getColumn();
  const 編集列名 = ヘッダー[col - 1] ? String(ヘッダー[col - 1]).trim() : '';

  if (!colMap.trigger.includes(編集列名)) return;

  const get = (名前) => 列[名前] ? sh.getRange(row, 列[名前]).getValue() : '';
  const set = (名前, v) => { if (列[名前] && v !== null && v !== undefined && v !== '') sh.getRange(row, 列[名前]).setValue(v); };

  let idType = null, id = null;

  if (編集列名 === colMap.url) {
    const itemId = アラジンURLからItemIdを抽出_(get(colMap.url));
    if (itemId) { idType = 'ItemId'; id = itemId; }
  } else if (colMap.isbn && 編集列名 === colMap.isbn) {
    const isbn = String(get(colMap.isbn)).replace(/[^\d]/g, '');
    if (isbn.length >= 10) { idType = 'ISBN13'; id = isbn; }
  }

  if (!idType || !id) return;

  const ttbKey = アラジンTTBキーを取得_();
  if (!ttbKey) {
    ss.toast('TTBキーが未設定です', '⚠ アラジンAPI', 5);
    return;
  }

  try {
    ss.toast('アラジンAPIに接続中...', '🔄 データ取得中', 10);
    const item = アラジンAPI呼び出し_(idType, id);
    if (!item) {
      ss.toast('商品が見つかりませんでした', '❌ 取得失敗', 4);
      return;
    }

    ss.toast('データをシートに書き込み中...', '🔄 データ取得中', 10);

    if (colMap.title)            set(colMap.title,            item.title || '');
    if (colMap.author)           set(colMap.author,           item.author || '');
    if (colMap.publisher)        set(colMap.publisher,        item.publisher || '');
    if (colMap.pubDate)          set(colMap.pubDate,          item.pubDate || '');
    if (colMap.price)            set(colMap.price,            item.priceSales || '');
    if (colMap.cover)            set(colMap.cover,            item.cover || '');
    if (colMap.additionalImages) set(colMap.additionalImages, item.cover || '');
    if (colMap.description)      set(colMap.description,      (item.fulldescription || item.description || '').slice(0, 2000));
    if (colMap.categoryName)     set(colMap.categoryName,     item.categoryName || '');
    if (colMap.itemId)           set(colMap.itemId,           String(item.itemId || ''));
    if (colMap.isbn && !get(colMap.isbn)) set(colMap.isbn,    item.isbn13 || item.isbn || '');

    if (colMap.url && !get(colMap.url) && item.itemId) {
      set(colMap.url, `https://www.aladin.co.kr/shop/wproduct.aspx?ItemId=${item.itemId}`);
    }

    ss.toast(item.title, '✅ データ取得完了', 5);
  } catch (err) {
    ss.toast(err.message, '❌ APIエラー', 6);
  }
}

// ============================================================
// 手動一括取得（選択行または全行）
// ============================================================
function アラジン一括取得() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const shName = sh.getName();
  const colMap = ALADIN_COLUMN_MAP[shName];

  if (!colMap) {
    ui.alert('このシートはアラジンAPI対象外です\n\n対象: 韓国書籍 / 韓国マンガ / 韓国音楽映像 / 韓国グッズ');
    return;
  }

  const ttbKey = アラジンTTBキーを取得_();
  if (!ttbKey) { ui.alert('TTBキーが未設定です\nメニュー → アラジンTTBキー設定 から登録してください'); return; }

  if (sh.getLastRow() < 2) { ui.alert('データがありません'); return; }

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  if (!列[colMap.url]) { ui.alert(`「${colMap.url}」列が見つかりません`); return; }

  const データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  let 成功 = 0, スキップ = 0, エラー = 0;
  const start = Date.now();

  データ.forEach((row, i) => {
    // 60秒タイムアウト対策
    if (Date.now() - start > 55000) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`⚠ タイムアウト手前で停止（${i}行処理済み）`, 'アラジン', 5);
      return;
    }

    const rowNum = i + 2;
    const url = String(row[列[colMap.url] - 1] || '').trim();
    if (!url) { スキップ++; return; }

    // すでにタイトルが入っていればスキップ
    if (colMap.title && 列[colMap.title] && row[列[colMap.title] - 1]) { スキップ++; return; }

    const itemId = アラジンURLからItemIdを抽出_(url);
    if (!itemId) { スキップ++; return; }

    try {
      const item = アラジンAPI呼び出し_('ItemId', itemId);
      if (!item) { スキップ++; return; }

      const setCell = (名前, v) => {
        if (列[名前] && v !== null && v !== undefined && v !== '') {
          sh.getRange(rowNum, 列[名前]).setValue(v);
        }
      };

      if (colMap.title)            setCell(colMap.title,            item.title || '');
      if (colMap.author)           setCell(colMap.author,           item.author || '');
      if (colMap.publisher)        setCell(colMap.publisher,        item.publisher || '');
      if (colMap.pubDate)          setCell(colMap.pubDate,          item.pubDate || '');
      if (colMap.price)            setCell(colMap.price,            item.priceSales || '');
      if (colMap.cover)            setCell(colMap.cover,            item.cover || '');
      if (colMap.additionalImages) setCell(colMap.additionalImages, item.cover || '');
      if (colMap.description)      setCell(colMap.description,      (item.fulldescription || item.description || '').slice(0, 2000));
      if (colMap.categoryName)     setCell(colMap.categoryName,     item.categoryName || '');
      if (colMap.itemId)           setCell(colMap.itemId,           String(item.itemId || ''));
      if (colMap.isbn)             setCell(colMap.isbn,             item.isbn13 || item.isbn || '');

      成功++;
      Utilities.sleep(300); // APIレート制限対策
    } catch (e) {
      Logger.log(`Row ${rowNum} エラー: ${e.message}`);
      エラー++;
    }
  });

  ui.alert(`✅ アラジン一括取得完了\n\n成功: ${成功}件\nスキップ: ${スキップ}件\nエラー: ${エラー}件`);
}

function アラジンテスト() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('韓国音楽映像');
  const url = sh.getRange('V8').getValue();
  const itemId = アラジンURLからItemIdを抽出_(url);
  const item = アラジンAPI呼び出し_('ItemId', itemId);
  
  Logger.log('subInfo: ' + JSON.stringify(item.subInfo));
}
function ロッククリア() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('__SELF_EDIT_LOCK__');
  props.deleteProperty('__SELF_EDIT_TS__');
  SpreadsheetApp.getActiveSpreadsheet().toast('ロッククリア完了', 'debug', 3);
}

function 権限テスト() {
  UrlFetchApp.fetch('https://www.google.com');
  SpreadsheetApp.getActiveSpreadsheet().toast('権限OK');
}
