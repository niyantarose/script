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
    + `&OptResult=authors,fulldescription`;

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
  '잡지': '雑誌', 'Magazine': '雑誌', '매거진': '雑誌',
  '문학': '書籍', '역사': '書籍',
  // グッズ
  'Gift': 'グッズ', 'Goods': 'グッズ', '굿즈': 'グッズ',
};

function アラジンジャンルを判定_(item) {
  if (!item) return 'その他';
  const cn = String(item.categoryName || '').toLowerCase();
  const ml = String(item.mallType || '').toLowerCase();
  const title = String(item.title || '').toLowerCase();
  const subtitle = String((item.subInfo && item.subInfo.subTitle) || '').toLowerCase();
  const combined = cn + ' ' + ml + ' ' + title + ' ' + subtitle;

  const hasMagazineKeyword = /magazine|잡지|매거진/i.test(combined);
  const hasMagazineIssuePattern = /\b20\d{2}[./-]\d{1,2}\b/.test(title) || /\b\d{4}年\s*\d{1,2}月\b/.test(title) || /\b\d{4}년\s*\d{1,2}월\b/.test(title);
  const hasMagazineBrand = /(elle|vogue|gq|esquire|allure|bazaar|harper'?s bazaar|marie claire|cosmopolitan|dazed|arena|w korea|ceci|1st look|cine21|maxim|singles|nylon|the star|star1|men'?s health)/i.test(title);

  if (hasMagazineKeyword || (hasMagazineIssuePattern && hasMagazineBrand)) return '雑誌';
  if (/music|음반|cd|lp/i.test(combined)) return '音楽映像';
  if (/dvd|bluray|blu-ray|블루레이|video/i.test(combined)) return '音楽映像';
  if (/comic|만화/i.test(combined)) return 'マンガ';
  if (/gift|goods|굿즈/i.test(combined)) return 'グッズ';
  if (/book|도서|novel|소설|essay|에세이/i.test(combined)) return '書籍';

  for (const [keyword, genre] of Object.entries(ALADIN_CATEGORY_MAP)) {
    if (combined.includes(keyword.toLowerCase())) return genre;
  }
  return '書籍'; // デフォルト
}

// ============================================================
// シート別列マッピング
// ============================================================
function 韓国マンガカテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const subtitle = String((source && source.subInfo && source.subInfo.subTitle) || '').trim();
  const combined = [categoryName, title, subtitle].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/essay|에세이/.test(combined)) return 'エッセイ';
  if (/novel|소설|라이트노벨|light\s*novel/.test(combined)) return '小説';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북/.test(combined)) return 'アートブック';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  return categoryName || '';
}

function 韓国音楽映像カテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const titleAndCategory = [title, categoryName].join(' ').toLowerCase();
  const combined = [title, categoryName, basicInfo, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/\bdvd\b|\[dvd\]|dvd\//.test(titleAndCategory)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(titleAndCategory)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(titleAndCategory)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(titleAndCategory)) return 'CD';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(titleAndCategory)) return 'CD';

  if (/\bdvd\b|\[dvd\]|dvd\//.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(combined)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(combined)) return 'CD';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(combined)) return 'CD';
  return categoryName || '';
}

function カテゴリ入力値を補正_(sheetName, source) {
  if (sheetName === '韓国マンガ') {
    return 韓国マンガカテゴリを補正_(source);
  }
  if (sheetName === '韓国音楽映像') {
    return 韓国音楽映像カテゴリを補正_(source);
  }
  return String(source && source.categoryName || '').trim();
}
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
    categoryName:     'カテゴリ',
    itemId:           'アラジン商品コード',
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
    categoryName:     'カテゴリ',
    itemId:           'アラジン商品コード',
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
    categoryName:     'カテゴリ',
    itemId:           'アラジン商品コード',
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
    itemId:           'アラジン商品コード',
  },
  '韓国雑誌': {
    trigger:          ['アラジンURL'],
    url:              'アラジンURL',
    title:            '原題商品名',
    pubDate:          '発売日',
    price:            '原価',
    cover:            'メイン画像URL',
    additionalImages: '追加画像URL',
    description:      '商品説明',
    itemId:           'アラジン商品コード',
  },
};

// ============================================================
// onEditから呼び出されるメイン処理
// ============================================================
function アラジン_onEdit(e) {
  const sh = e.range.getSheet();
  const shName = sh.getName();
  const colMap = ALADIN_COLUMN_MAP[shName];
  if (!colMap) return;

  const row = e.range.getRow();
  if (row < 2) return;
  if (自己更新中か_()) return;

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  const col = e.range.getColumn();
  const 編集列名 = ヘッダー[col - 1] ? String(ヘッダー[col - 1]).trim() : '';

  // トリガー列かチェック
  if (!colMap.trigger.includes(編集列名)) return;

  const get = (名前) => 列[名前] ? sh.getRange(row, 列[名前]).getValue() : '';
  const set = (名前, v) => { if (列[名前] && v !== null && v !== undefined && v !== '') sh.getRange(row, 列[名前]).setValue(v); };

  // ItemId またはISBN を特定
  let idType = null, id = null;

  if (編集列名 === colMap.url) {
    const url = get(colMap.url);
    const itemId = アラジンURLからItemIdを抽出_(url);
    if (itemId) { idType = 'ItemId'; id = itemId; }
  } else if (colMap.isbn && 編集列名 === colMap.isbn) {
    const isbn = String(get(colMap.isbn)).replace(/[^\d]/g, '');
    if (isbn.length >= 10) { idType = 'ISBN13'; id = isbn; }
  }

  if (!idType || !id) return;

  // API呼び出し（時間がかかるのでロック不要、自己更新フラグだけ立てる）
  const ttbKey = アラジンTTBキーを取得_();
  if (!ttbKey) {
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠ TTBキーが未設定です（メニュー → アラジンTTBキー設定）', 'アラジンAPI', 5);
    return;
  }

  自己更新を開始_();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('⏳ アラジンAPIから取得中...', 'アラジン', 3);
    const item = アラジンAPI呼び出し_(idType, id);
    if (!item) {
      SpreadsheetApp.getActiveSpreadsheet().toast('❌ 商品が見つかりませんでした', 'アラジン', 4);
      return;
    }

    // 各列にセット（既入力の値は上書きしない）
    if (colMap.title)            set(colMap.title,            item.title || '');
    if (colMap.author)           set(colMap.author,           item.author || '');
    if (colMap.publisher)        set(colMap.publisher,        item.publisher || '');
    if (colMap.pubDate)          set(colMap.pubDate,          item.pubDate || '');
    if (colMap.price)            set(colMap.price,            item.priceSales || '');
    if (colMap.cover)            set(colMap.cover,            item.cover || '');
    if (colMap.additionalImages) set(colMap.additionalImages, item.cover || '');
    if (colMap.description)      set(colMap.description,      (item.fulldescription || item.description || '').slice(0, 2000));
    if (colMap.categoryName)     set(colMap.categoryName,     カテゴリ入力値を補正_(shName, item));
    if (colMap.itemId)           set(colMap.itemId,           String(item.itemId || ''));
    if (colMap.isbn && !get(colMap.isbn)) set(colMap.isbn,    item.isbn13 || item.isbn || '');

    // URLがまだ空なら補完
    if (colMap.url && !get(colMap.url) && item.itemId) {
      set(colMap.url, `https://www.aladin.co.kr/shop/wproduct.aspx?ItemId=${item.itemId}`);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ 取得完了: ${item.title}`, 'アラジン', 4);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`❌ API エラー: ${err.message}`, 'アラジン', 6);
    Logger.log('アラジン_onEdit エラー: ' + err.message);
  } finally {
    自己更新を終了_();
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
    ui.alert('このシートはアラジンAPI対象外です\n\n対象: 韓国書籍 / 韓国雑誌 / 韓国マンガ / 韓国音楽映像 / 韓国グッズ');
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
      if (colMap.categoryName)     setCell(colMap.categoryName,     カテゴリ入力値を補正_(shName, item));
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




