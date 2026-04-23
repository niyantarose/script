// アラジンウェブアプリ.gs
// Chrome拡張機能からのデータを受け取ってシートに書き込む
// ウェブアプリとしてデプロイ必要（全員アクセス可）
// 共通マスターファイルに配置して、カテゴリごとに各スプレッドシートへ振り分ける

const ALADIN_SPREADSHEET_DESTINATION_MAP = {
  '韓国書籍': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国音楽映像': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国雑誌': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国マンガ': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1fItni5nPV_tLshtKQBXe8_Ag965G-nBn25JPvgg6BU4/edit?gid=1100321429#gid=1100321429',
    spreadsheetLabel: '商品登録よろしく_韓国マンガ'
  },
  '韓国グッズ': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1BwNNvx0PakFzugprDJKmmZ5tjDTZijKZ8VH-anTuwDU/edit?gid=1562610398#gid=1562610398',
    spreadsheetLabel: '商品登録よろしく_韓国グッズ'
  }
};

const ALADIN_ALL_TARGET_SHEETS = ['韓国書籍', '韓国雑誌', '韓国マンガ', '韓国音楽映像', '韓国グッズ'];
const ALADIN_GENRE_SHEET_MAP = {
  '書籍': ['韓国書籍'],
  '雑誌': ['韓国雑誌'],
  'マンガ': ['韓国マンガ'],
  '音楽映像': ['韓国音楽映像'],
  'グッズ': ['韓国グッズ']
};

const ALADIN_APPEND_VALUE_HEADERS = [
  'アラジン商品コード', 'サイト商品コード', 'ItemId',
  'アラジンURL', '購入URL', 'リンク',
  'ISBN', 'JANコード',
  '商品名(日本語)', '商品名（日本語）', '日本語タイトル',
  '商品名(原題)', '商品名（原題）', '作品名（原題）', '原題タイトル', '原題商品タイトル',
  '雑誌名', '著者', '作者', '出版社', 'アーティスト名', 'アーティスト', '事務所/レーベル',
  '発売日', '原価', 'メイン画像URL', 'メイン画像', '追加画像URL', '追加画像',
  '商品説明', '備考', 'カテゴリ', 'ジャンル分類', '特典メモ', '付録情報',
  '表紙情報', '年', '月', '号数'
];

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    if (data.action === 'updateAladinData') {
      const result = アラジンデータを更新_(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function アラジンデータを更新_(data) {
  const itemId = data && data.itemId;
  const genre = 補正ジャンルを判定_(data);
  const normalizedData = Object.assign({}, data, { genre: genre });
  if (!itemId) return { success: false, error: 'itemId missing' };

  const targetSheets = 保存先シート候補を取得_(genre);
  if (!targetSheets.length) {
    return { success: false, error: 保存先未設定メッセージを組み立て_(genre) };
  }

  for (const target of targetSheets) {
    const sh = target.sheet;
    if (!sh || sh.getLastRow() < 2) continue;

    const context = シートコンテキストを取得_(sh);
    if (!context.idCol) continue;

    const row = ItemId一致行を探す_(sh, context.idCol, itemId);
    if (!row) continue;

    アラジン行へ書き込む_(sh, context, row, normalizedData);
    return {
      success: true,
      sheet: target.sheetName,
      spreadsheetId: target.spreadsheetId,
      spreadsheetName: target.spreadsheetName,
      row,
      resolvedGenre: genre,
      mode: 'updated'
    };
  }

  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const createTarget = targetSheets.find(target => preferredSheets.includes(target.sheetName)) || targetSheets[0];
  if (!createTarget || !createTarget.sheet) {
    return { success: false, error: '対象シートが見つかりませんでした' };
  }

  const createSheet = createTarget.sheet;
  const context = シートコンテキストを取得_(createSheet);
  const row = 追加対象行を決める_(createSheet, context);

  if (row > createSheet.getMaxRows()) {
    createSheet.insertRowsAfter(createSheet.getMaxRows(), row - createSheet.getMaxRows());
  }

  アラジン行へ書き込む_(createSheet, context, row, normalizedData);
  return {
    success: true,
    sheet: createTarget.sheetName,
    spreadsheetId: createTarget.spreadsheetId,
    spreadsheetName: createTarget.spreadsheetName,
    row,
    resolvedGenre: genre,
    mode: 'created'
  };
}

function 補正ジャンルを判定_(data) {
  const requestedGenre = String(data && data.genre || '').trim();
  const categoryName = String(data && data.categoryName || '').toLowerCase();
  const mallType = String(data && data.mallType || '').toLowerCase();
  const title = String(data && data.title || '').toLowerCase();
  const combined = [requestedGenre, categoryName, mallType, title].join(' ');

  const hasMagazineKeyword = /magazine|잡지|매거진|雑誌/i.test(combined);
  const hasMagazineIssuePattern = /\b20\d{2}[./-]\d{1,2}\b/.test(title) || /\b\d{4}年\s*\d{1,2}月\b/.test(title) || /\b\d{4}년\s*\d{1,2}월\b/.test(title);
  const hasMagazineBrand = /(elle|vogue|gq|esquire|allure|bazaar|harper'?s bazaar|marie claire|cosmopolitan|dazed|arena|w korea|ceci|1st look|cine21|maxim|singles|nylon|the star|star1|men'?s health)/i.test(title);

  if (requestedGenre === '雑誌' || hasMagazineKeyword || (hasMagazineIssuePattern && hasMagazineBrand)) {
    return '雑誌';
  }

  return requestedGenre || '書籍';
}

const MAGAZINE_NAME_PATTERNS_ = [
  { name: 'DAZED & CONFUSED', pattern: /dazed\s*&\s*confused/i },
  { name: "HARPER'S BAZAAR", pattern: /harper'?s\s*bazaar/i },
  { name: 'BAZAAR', pattern: /\bbazaar\b/i },
  { name: 'MARIE CLAIRE', pattern: /marie\s*claire/i },
  { name: 'COSMOPOLITAN', pattern: /cosmopolitan/i },
  { name: 'ESQUIRE', pattern: /esquire/i },
  { name: 'ALLURE', pattern: /allure/i },
  { name: 'VOGUE', pattern: /vogue/i },
  { name: 'ELLE', pattern: /elle/i },
  { name: 'GQ', pattern: /\bgq\b/i },
  { name: 'ARENA', pattern: /arena/i },
  { name: 'W KOREA', pattern: /\bw\s*korea\b/i },
  { name: 'CECI', pattern: /\bceci\b/i },
  { name: '1ST LOOK', pattern: /1st\s*look/i },
  { name: 'CINE21', pattern: /cine\s*21/i },
  { name: 'MAXIM', pattern: /maxim/i },
  { name: 'SINGLES', pattern: /singles/i },
  { name: 'NYLON', pattern: /nylon/i },
  { name: 'THE STAR', pattern: /the\s*star/i },
  { name: 'STAR1', pattern: /star\s*1|star1/i },
  { name: "MEN'S HEALTH", pattern: /men'?s\s*health/i }
];

function 雑誌名を抽出_(data) {
  const explicitName = String(data && data.magazineName || '').trim();
  if (explicitName) return explicitName;

  const title = String(data && data.title || '').trim();
  const categoryName = String(data && data.categoryName || '').trim();
  const combined = `${title} ${categoryName}`;

  for (const entry of MAGAZINE_NAME_PATTERNS_) {
    if (entry.pattern.test(combined)) {
      return entry.name;
    }
  }

  const cleaned = title
    .replace(/\b20\d{2}[./-]\d{1,2}\b.*$/i, '')
    .replace(/\b\d{4}년\s*\d{1,2}월.*$/i, '')
    .replace(/\b\d{4}年\s*\d{1,2}月.*$/i, '')
    .replace(/\s+(KOREA|코리아)\b/ig, '')
    .replace(/[<＜【\[(].*$/, '')
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned || title;
}

function 雑誌号情報を抽出_(data) {
  const title = String(data && data.title || '').trim();
  const basicInfo = String(data && data.basicInfo || '').trim();
  const description = String(data && data.description || '').trim();
  const magazineName = 雑誌名を抽出_(data);
  const combined = [title, basicInfo, description].filter(Boolean).join(' / ');

  let year = '';
  let month = '';
  let issue = '';

  const yearMonthPatterns = [
    /\b(20\d{2})[./-]\s*(\d{1,2})\b/i,
    /\b(20\d{2})年\s*(\d{1,2})月\b/i,
    /\b(20\d{2})년\s*(\d{1,2})월\b/i
  ];
  for (const pattern of yearMonthPatterns) {
    const match = combined.match(pattern);
    if (match) {
      year = match[1];
      month = String(Number(match[2]));
      break;
    }
  }

  let rest = title;
  if (magazineName) {
    const escapedMagazine = magazineName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    rest = rest.replace(new RegExp(escapedMagazine, 'ig'), ' ');
  }
  rest = rest
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/ig, ' ')
    .replace(/\b20\d{2}[./-]\d{1,2}\b/ig, ' ')
    .replace(/\b\d{4}年\s*\d{1,2}月\b/ig, ' ')
    .replace(/\b\d{4}년\s*\d{1,2}월\b/ig, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const issuePatterns = [
    /\b(?:VOL\.?|NO\.?|ISSUE|#)\s*([1-9]\d{0,3})\b/i,
    /\b([1-9]\d{0,3})\s*호\b/i,
    /\b([1-9]\d{0,2})\b/
  ];
  for (const pattern of issuePatterns) {
    const match = rest.match(pattern);
    if (!match) continue;
    const candidate = String(match[1] || '').trim();
    if (!candidate) continue;
    if (year && candidate === year) continue;
    if (month && candidate === month) continue;
    if (Number(candidate) > 500) continue;
    issue = candidate;
    break;
  }

  return { year, month, issue };
}

function 韓国マンガカテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const combined = [categoryName, title, basicInfo, description].join(' ').toLowerCase();

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
  return 'まんが';
}

function 韓国書籍カテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const combined = [categoryName, title, basicInfo, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트ラック/.test(combined)) return 'OST';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/참고서|수험서|문제집|기출|모의고사|수능|내신|검정고시|자격증|공무원|고시|임용|편입|leet|meet|deet|psat|ncs|toeic|toefl|ielts|teps|jlpt|hsk|topik|参考書|問題集|過去問|受験|資格/.test(combined)) return '参考書';
  if (/교재|학습지|학습서|워크북|work\s*book|workbook|text\s*book|textbook|student\s*book|activity\s*book|teacher'?s\s*book|course\s*book|教材|教科書/.test(combined)) return '教材';
  if (/essay|에세이|산문/.test(combined)) return 'エッセイ';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북|예체능|미술|디자인|사진/.test(combined)) return 'アートブック';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  if (/novel|소설|라이트노벨|light\s*novel|문학|시집/.test(combined)) return '小説';
  return '';
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
  if (sheetName === '韓国書籍') {
    return 韓国書籍カテゴリを補正_(source);
  }
  if (sheetName === '韓国マンガ') {
    return 韓国マンガカテゴリを補正_(source);
  }
  if (sheetName === '韓国音楽映像') {
    return 韓国音楽映像カテゴリを補正_(source);
  }
  return String(source && source.categoryName || '').trim();
}
function 保存先シート候補を取得_(genre) {
  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const targetSheetNames = preferredSheets.length ? preferredSheets : ALADIN_ALL_TARGET_SHEETS;
  const spreadsheetCache = {};
  const targets = [];

  targetSheetNames.forEach(sheetName => {
    const target = 保存先シートを取得_(sheetName, spreadsheetCache);
    if (target) targets.push(target);
  });

  return targets;
}

function 保存先シートを取得_(sheetName, spreadsheetCache) {
  const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
  const spreadsheetId = スプレッドシートIDを抽出_(destination.spreadsheetIdOrUrl || '');
  if (!spreadsheetId) return null;

  let ss = spreadsheetCache[spreadsheetId];
  if (!ss) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      throw new Error(`保存先を開けませんでした: ${sheetName} (${destination.spreadsheetLabel || spreadsheetId}) / ${error.message}`);
    }
    spreadsheetCache[spreadsheetId] = ss;
  }

  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    throw new Error(`保存先シートが見つかりません: ${sheetName} (${destination.spreadsheetLabel || ss.getName()})`);
  }

  return {
    sheetName,
    sheet: sh,
    spreadsheetId,
    spreadsheetName: ss.getName()
  };
}

function 保存先未設定メッセージを組み立て_(genre) {
  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const targetSheetNames = preferredSheets.length ? preferredSheets : ALADIN_ALL_TARGET_SHEETS;
  const missingLabels = targetSheetNames
    .filter(sheetName => {
      const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
      return !スプレッドシートIDを抽出_(destination.spreadsheetIdOrUrl || '');
    })
    .map(sheetName => {
      const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
      return `${sheetName} -> ${destination.spreadsheetLabel || '未設定'}`;
    });

  if (!missingLabels.length) {
    return '保存先シートが見つかりませんでした';
  }

  return `保存先スプレッドシート未設定: ${missingLabels.join(', ')}`;
}

function スプレッドシートIDを抽出_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : text;
}

function シートコンテキストを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  headers.forEach((header, index) => {
    if (header) 列[String(header).trim()] = index + 1;
  });

  const fallbackMap = {
    '韓国書籍': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'ISBN',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      author: ['作者', '著者'],
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国マンガ': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'ISBN',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      author: ['作者', '著者'],
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国音楽映像': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'JANコード',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      author: ['アーティスト名', 'アーティスト'],
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国グッズ': {
      url: ['リンク', '購入URL', 'アラジンURL'],
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国雑誌': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      title: ['原題商品タイトル', '原題商品名', '原題タイトル', '商品名（原題）', '商品名(原題)'],
      magazineName: '雑誌名',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    }
  };

  const columnMapSource = (typeof ALADIN_COLUMN_MAP !== 'undefined' && ALADIN_COLUMN_MAP[sh.getName()])
    ? ALADIN_COLUMN_MAP[sh.getName()]
    : fallbackMap[sh.getName()] || {};

  return {
    headers,
    列,
    colMap: columnMapSource,
    idCol: 最初に見つかった列番号を返す_(列, ['サイト商品コード', 'アラジン商品コード', 'ItemId'])
  };
}

function 最初に見つかった列番号を返す_(列, headerNames) {
  const names = Array.isArray(headerNames) ? headerNames : [headerNames];
  for (const name of names) {
    if (name && 列[name]) return 列[name];
  }
  return null;
}

function 最初に見つかったヘッダー名を返す_(列, headerNames) {
  const names = Array.isArray(headerNames) ? headerNames : [headerNames];
  return names.find(name => name && 列[name]) || null;
}

function 追加対象行を決める_(sh, context) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), 1);
  if (lastRow < 2) return 2;

  const headers = context && Array.isArray(context.headers)
    ? context.headers.map(header => String(header || '').trim())
    : sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(header => String(header || '').trim());
  const appendTargetIndexes = headers
    .map((header, index) => ALADIN_APPEND_VALUE_HEADERS.includes(header) ? index : -1)
    .filter(index => index >= 0);
  const indexes = appendTargetIndexes.length
    ? appendTargetIndexes
    : Array.from({ length: lastColumn }, (_, index) => index);

  const values = sh.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const hasVisibleValue = indexes.some(index => {
      const value = String(values[i][index] || '').trim();
      return value !== '' && value !== 'FALSE';
    });
    if (hasVisibleValue) {
      return i + 3;
    }
  }

  return 2;
}

function ItemId一致行を探す_(sh, idCol, itemId) {
  if (!idCol || sh.getLastRow() < 2) return null;

  const values = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(itemId).trim()) {
      return i + 2;
    }
  }
  return null;
}

function アラジン行へ書き込む_(sh, context, row, data) {
  const { 列, colMap } = context;
  const setByHeader = (headerNames, value) => {
    const headerName = 最初に見つかったヘッダー名を返す_(列, headerNames);
    if (!headerName) return;
    if (value === null || value === undefined || value === '') return;
    sh.getRange(row, 列[headerName]).setValue(value);
  };
  const setIfBlank = (headerNames, value) => {
    const headerName = 最初に見つかったヘッダー名を返す_(列, headerNames);
    if (!headerName) return;
    if (value === null || value === undefined || value === '') return;
    const range = sh.getRange(row, 列[headerName]);
    if (String(range.getDisplayValue() || '').trim()) return;
    range.setValue(value);
  };

  const basicInfoCol = 列['基本情報'] || null;
  const mergedDescription = 商品説明テキストを組み立てる_(data);

  setByHeader(colMap.url, data.pageUrl || '');
  setByHeader(colMap.isbn, data.isbn13 || '');
  setByHeader(colMap.title, data.title || '');
  setByHeader(colMap.author, data.author || '');
  setByHeader(colMap.publisher, data.publisher || '');
  setByHeader(colMap.pubDate, data.pubDate || '');
  setByHeader(colMap.price, data.priceSales || '');
  setByHeader(colMap.cover, data.cover || '');
  setByHeader(colMap.additionalImages, data.additionalImages || '');
  setByHeader(colMap.description, mergedDescription);
  setByHeader(colMap.categoryName, カテゴリ入力値を補正_(sh.getName(), data));
  setByHeader(colMap.itemId, String(data.itemId || ''));

  if (sh.getName() === '韓国雑誌') {
    const magazineName = 雑誌名を抽出_(data);
    const issueInfo = 雑誌号情報を抽出_(data);
    setByHeader(colMap.magazineName || '雑誌名', magazineName);
    setByHeader('原題タイトル', magazineName || data.title || '');
    setByHeader('原題商品名', data.title || '');
    setIfBlank('年', issueInfo.year);
    setIfBlank('月', issueInfo.month);
    setIfBlank('号数', issueInfo.issue);
  }

  if (basicInfoCol && data.basicInfo) {
    const existing = sh.getRange(row, basicInfoCol).getValue();
    if (!existing) {
      sh.getRange(row, basicInfoCol).setValue(String(data.basicInfo).slice(0, 500));
    }
  }

  行表示を整える_(sh, row);
}

function 一行テキストへ整形_(value) {
  return String(value || '')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .join(' / ');
}

function 商品説明テキストを組み立てる_(data) {
  const description = 一行テキストへ整形_(data.description || '');
  const basicInfo = 一行テキストへ整形_(data.basicInfo || '');
  const parts = [];

  if (basicInfo) {
    parts.push(`[基本情報] ${basicInfo}`);
  }
  if (description) {
    parts.push(description);
  }

  return parts.join(' / ').slice(0, 4000);
}

function 行表示を整える_(sh, row) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const range = sh.getRange(row, 1, 1, lastColumn);
  range.setWrap(false);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  range.setVerticalAlignment('middle');
  sh.setRowHeight(row, 21);
}





