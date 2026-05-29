/**
 * books.com.tw（博客來）Chrome拡張機能 用
 * スプレッドシート追記GAS
 *
 * 対応:
 * - 台湾書籍（コミック/小説/設定集）
 * - 台湾グッズ
 * - 台湾雑誌
 *
 * 雑誌まわりの強化:
 * - 共通マスター参照
 * - 別名1〜n による雑誌照合
 * - 未登録雑誌の候補蓄積
 * - 言語接頭辞付き親コード生成
 */

const SPREADSHEET_ID = '1OSoZnNTMHrH5YgU-j7zwOsyQsjS6bcwjfZxc3kzeKWM';
const MASTER_SPREADSHEET_ID = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';

const COMIC_BOOK_SHEET_NAME = '台湾まんが';
const OTHER_BOOK_SHEET_NAME = '台湾書籍その他';
const GOODS_SHEET_NAME = '台湾グッズ';
const MAGAZINE_SHEET_NAME = '台湾雑誌';

const MAGAZINE_MASTER_SHEET = '雑誌マスター（共通）';
const MAGAZINE_MASTER_CANDIDATE_SHEET = '雑誌マスター候補（共通）';
const MAGAZINE_EDITION_RULE_SHEET = '版種ルール（雑誌共通）';
const MAGAZINE_TYPE_RULE_SHEET = 'タイプルール（雑誌共通）';
const JAPANESE_TITLE_DICTIONARY_SHEET_NAME = '日本語タイトル辞書';
/**
 * 共通マスター（MASTER_SPREADSHEET_ID）内の「作品タイトル」蓄積シート。
 * ブラウザURLの gid= と同じ数値。1行目ヘッダは「日本語タイトル辞書」と同一推奨:
 * 言語 / カテゴリ / 原題タイトル / 日本語タイトル / 著者（任意） / 別名（任意・区切りは改行・;・, など）
 */
const WORK_TITLE_MASTER_SHEET_GID = 1630833397;

const UNKNOWN_MAGAZINE_WRITE_MODE = 'candidate';
const ALLOW_PROVISIONAL_MAGAZINE_CODE = true;
const MAGAZINE_PARENT_CODE_WITH_LANGUAGE_PREFIX = true;
const JAPANESE_TITLE_NOT_FOUND_LABEL = '登録なし';
const JAPANESE_TITLE_LOOKUP_FAILED_LABEL = '照会失敗';
const JAPANESE_TITLE_NOT_LOOKED_UP_LABEL = '未照会';
const JAPANESE_TITLE_NOT_FOUND_ALL_SOURCES_LABEL = '登録なし(全サイト)';
const FAST_DUPLICATE_CHECK_ITEM_LIMIT = 3;
const LOOKUP_WRITEBACK_HEADERS = [
  '日本語タイトル',
  '作品名（日本語）',
  '日本語タイトル照会結果',
  '日本語タイトル照会元',
  '日本語タイトル照会ログ',
  '検索用正規化タイトル'
];
const BUILTIN_TITLE_ALIAS_DICTIONARY = {
  '台湾|まんが|排球少年': 'ハイキュー!!',
  '台湾|グッズ|排球少年': 'ハイキュー!!',
  '台湾|まんが|我的英雄學院': '僕のヒーローアカデミア',
  '台湾|グッズ|我的英雄學院': '僕のヒーローアカデミア',
  '台湾|まんが|咒術迴戰': '呪術廻戦',
  '台湾|グッズ|咒術迴戰': '呪術廻戦',
  '台湾|まんが|葬送的芙莉蓮': '葬送のフリーレン',
  '台湾|書籍|葬送的芙莉蓮': '葬送のフリーレン',
  '台湾|まんが|我獨自升級': '俺だけレベルアップな件',
  '台湾|書籍|我獨自升級': '俺だけレベルアップな件',
  '台湾|グッズ|我獨自升級': '俺だけレベルアップな件',
  /** Roses and Champagne（英題・MU）— GAS から MU が 403 のときのフォールバック */
  '台湾|まんが|薔薇與香檳': '薔薇とシャンパン',
  '台湾|グッズ|薔薇與香檳': '薔薇とシャンパン',
  '台湾|まんが|rosesandchampagne': '薔薇とシャンパン',
  '台湾|グッズ|rosesandchampagne': '薔薇とシャンパン',
  /** Saving My Sweetheart（MU associated に日本語あり） */
  '台湾|書籍|只想守護溫柔的你': '優しいあなたを守る方法',
  '台湾|まんが|只想守護溫柔的你': '優しいあなたを守る方法',
  '台湾|グッズ|只想守護溫柔的你': '優しいあなたを守る方法',
  '台湾|書籍|savingmysweetheart': '優しいあなたを守る方法',
  '台湾|まんが|savingmysweetheart': '優しいあなたを守る方法'
};
const TITLE_LOOKUP_PROVIDERS = {
  workTitleMaster: {
    label: '作品タイトルマスター（共通シート）',
    role: 'マスターSP（gid）に登録した別名・原題から日本語正式題への変換',
    priority: 'highest',
    implemented: true
  },
  titleAliasDictionary: {
    label: '自社タイトル別名辞書',
    role: '内蔵辞書＋「日本語タイトル辞書」シートから日本語タイトルへの確定変換',
    priority: 'highest',
    implemented: true
  },
  chilchil: {
    label: 'ちるちる',
    role: 'BL漫画・BL小説の日本語タイトル確認',
    itemTypes: ['bl_manga', 'bl_novel'],
    implemented: true,
    failSoft: true
  },
  mangaUpdates: {
    label: 'MangaUpdates',
    role: '漫画・manhwa・manhuaの原題/英題/別名/シリーズ確認',
    itemTypes: ['manga', 'bl_manga', 'goods', 'light_novel'],
    implemented: true,
    failSoft: true,
    /** GoogleのIPはMU側で恒久ブロックされるため、GASでは実行せずブラウザ拡張に委譲 */
    delegatedToExtension: true
  },
  aniList: {
    label: 'AniList',
    role: 'アニメ・漫画・メディアミックス作品のタイトル確認',
    itemTypes: ['manga', 'bl_manga', 'goods', 'light_novel'],
    implemented: true,
    failSoft: true
  },
  myAnimeList: {
    label: 'MyAnimeList',
    role: 'アニメ・漫画系タイトル補助',
    itemTypes: ['manga', 'goods', 'light_novel'],
    optional: true,
    implemented: false
  },
  mangaDex: {
    label: 'MangaDex',
    role: '多言語タイトル・別名補助',
    itemTypes: ['manga', 'goods'],
    optional: true,
    implemented: false
  },
  bookWalker: {
    label: 'BOOK☆WALKER',
    role: '日本語漫画・ラノベ・電子書籍タイトル確認',
    itemTypes: ['manga', 'light_novel', 'novel_book', 'bl_manga'],
    implemented: true,
    failSoft: true
  },
  amazon: {
    label: 'Amazon JP',
    role: '書籍・ラノベ・グッズ・一般書籍の日本語タイトル補助',
    itemTypes: ['manga', 'bl_manga', 'light_novel', 'novel_book', 'goods'],
    implemented: true,
    failSoft: true
  },
  googleBooks: {
    label: 'Google Books API',
    role: '書籍・雑誌メタデータ補助',
    itemTypes: ['novel_book', 'light_novel', 'magazine'],
    implemented: false
  },
  openBD: {
    label: 'openBD',
    role: '日本のISBN書誌確認',
    itemTypes: ['novel_book', 'light_novel', 'manga'],
    implemented: false
  },
  ndlSearch: {
    label: '国立国会図書館サーチ',
    role: '日本の書誌情報確認',
    itemTypes: ['novel_book', 'light_novel', 'manga'],
    optional: true,
    implemented: false
  },
  magazineMaster: {
    label: '自社雑誌マスター',
    role: '雑誌名・年月・号数の正規化',
    itemTypes: ['magazine'],
    implemented: false
  }
};
const PROVIDER_ORDER_BY_ITEM_TYPE = {
  manga: ['workTitleMaster', 'titleAliasDictionary', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon', 'myAnimeList', 'mangaDex'],
  bl_manga: ['workTitleMaster', 'titleAliasDictionary', 'chilchil', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon'],
  light_novel: ['workTitleMaster', 'titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'aniList', 'mangaUpdates'],
  novel_book: ['workTitleMaster', 'titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'ndlSearch'],
  goods: ['workTitleMaster', 'titleAliasDictionary', 'aniList', 'mangaUpdates', 'bookWalker', 'amazon'],
  magazine: ['workTitleMaster', 'magazineMaster', 'googleBooks', 'amazon', 'bookWalker'],
  music_video: ['workTitleMaster', 'titleAliasDictionary', 'amazon'],
  unknown: ['workTitleMaster', 'titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks']
};
const PROVIDER_ID_ALIASES = {
  worktitlemaster: 'workTitleMaster',
  work_title_master: 'workTitleMaster',
  workmaster: 'workTitleMaster',
  作品マスター: 'workTitleMaster',
  dictionary: 'titleAliasDictionary',
  titleDictionary: 'titleAliasDictionary',
  title_alias_dictionary: 'titleAliasDictionary',
  mangaupdates: 'mangaUpdates',
  manga_updates: 'mangaUpdates',
  anilist: 'aniList',
  bookwalker: 'bookWalker',
  book_walker: 'bookWalker',
  magazine_master: 'magazineMaster',
  myanimelist: 'myAnimeList',
  my_anime_list: 'myAnimeList',
  mangadex: 'mangaDex',
  manga_dex: 'mangaDex',
  googlebooks: 'googleBooks',
  google_books: 'googleBooks',
  ndl: 'ndlSearch',
  ndl_search: 'ndlSearch'
};

let _masterSpreadsheetCache = null;
let _japaneseTitleDictionaryCache = null;
let _workTitleMasterCache = null;

function getMasterSpreadsheet_() {
  if (_masterSpreadsheetCache) return _masterSpreadsheetCache;
  _masterSpreadsheetCache = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  return _masterSpreadsheetCache;
}

function buildJapaneseTitleDictionaryMapFromSheet_(sheet) {
  const byKey = {};
  if (!sheet) return byKey;

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return byKey;

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headerMap = {};
  values[0].forEach(function(header, idx) {
    const normalized = normalizeHeader_(header);
    if (normalized) headerMap[normalized] = idx;
  });

  const languageIdx = headerMap[normalizeHeader_('言語')];
  const categoryIdx = headerMap[normalizeHeader_('カテゴリ')];
  const originalTitleIdx = headerMap[normalizeHeader_('原題タイトル')];
  const japaneseTitleIdx = headerMap[normalizeHeader_('日本語タイトル')];
  const authorIdx = headerMap[normalizeHeader_('著者')];
  const aliasIdx = headerMap[normalizeHeader_('別名')];

  if (languageIdx === undefined || categoryIdx === undefined || originalTitleIdx === undefined || japaneseTitleIdx === undefined) {
    return byKey;
  }

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    const language = normalizeDictionaryLanguage_(row[languageIdx] || '');
    const category = normalizeDictionaryCategory_(row[categoryIdx] || '');
    const japaneseTitle = String(row[japaneseTitleIdx] || '').trim();
    const authorKey = authorIdx !== undefined ? normalizeDictionaryAuthorKey_(row[authorIdx] || '') : '';
    const titleCandidates = [row[originalTitleIdx] || ''].concat(parseDictionaryAliases_(aliasIdx >= 0 ? row[aliasIdx] : ''));

    if (!language || !category || !japaneseTitle) continue;

    titleCandidates.forEach(function(titleCandidate, candidateIndex) {
      const titleKey = normalizeDictionaryTitleKey_(titleCandidate);
      if (!titleKey) return;

      const key = buildJapaneseTitleDictionaryKey_(language, category, titleKey);
      if (!byKey[key]) byKey[key] = [];
      byKey[key].push({
        japaneseTitle: japaneseTitle,
        authorKey: authorKey,
        isAlias: candidateIndex > 0,
      });
    });
  }

  return byKey;
}

function getJapaneseTitleDictionaryIndex_() {
  if (_japaneseTitleDictionaryCache) return _japaneseTitleDictionaryCache;

  const master = getMasterSpreadsheet_();
  const sheet = master.getSheetByName(JAPANESE_TITLE_DICTIONARY_SHEET_NAME);
  const byKey = buildJapaneseTitleDictionaryMapFromSheet_(sheet);

  _japaneseTitleDictionaryCache = {
    available: !!sheet,
    byKey: byKey,
  };
  return _japaneseTitleDictionaryCache;
}

function getWorkTitleMasterIndex_() {
  if (_workTitleMasterCache) return _workTitleMasterCache;

  var sheet = null;
  try {
    sheet = getMasterSpreadsheet_().getSheetById(WORK_TITLE_MASTER_SHEET_GID);
  } catch (e) {
    sheet = null;
  }

  const byKey = buildJapaneseTitleDictionaryMapFromSheet_(sheet);
  _workTitleMasterCache = {
    available: !!sheet,
    byKey: byKey,
  };
  return _workTitleMasterCache;
}

function findTitleInDictionaryByKeyMap_(byKeyMap, language, category, originalTitle, author, originalProductTitle) {
  const normalizedLanguage = normalizeDictionaryLanguage_(language);
  const normalizedCategory = normalizeDictionaryCategory_(category);
  const titleKeys = buildDictionaryTitleKeys_(originalTitle, normalizedLanguage, normalizedCategory, originalProductTitle);
  if (!normalizedLanguage || !normalizedCategory || !titleKeys.length || !byKeyMap) return '';

  const authorKey = normalizeDictionaryAuthorKey_(author);
  const seen = {};
  const candidates = [];
  titleKeys.forEach(function(titleKey) {
    const bucket = byKeyMap[buildJapaneseTitleDictionaryKey_(normalizedLanguage, normalizedCategory, titleKey)] || [];
    bucket.forEach(function(candidate) {
      const dedupeKey = [
        String((candidate && candidate.japaneseTitle) || '').trim(),
        String((candidate && candidate.authorKey) || '').trim(),
        candidate && candidate.isAlias ? '1' : '0',
      ].join('::');
      if (!dedupeKey || seen[dedupeKey]) return;
      seen[dedupeKey] = true;
      candidates.push(candidate);
    });
  });
  if (!candidates.length) {
    return '';
  }

  const ranked = candidates.slice().sort(function(left, right) {
    const leftScore = (authorKey && left.authorKey === authorKey ? 100 : 0) + (left.isAlias ? 0 : 10);
    const rightScore = (authorKey && right.authorKey === authorKey ? 100 : 0) + (right.isAlias ? 0 : 10);
    return rightScore - leftScore;
  });

  return String((ranked[0] && ranked[0].japaneseTitle) || '').trim();
}

function buildJapaneseTitleDictionaryKey_(language, category, titleKey) {
  return [language, category, titleKey].join('|');
}

function normalizeDictionaryLanguage_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^(?:tw|台湾|台灣|繁體|繁体|繁體中文|繁体中文)$/i.test(text) || /繁體|繁体|台灣|台湾/.test(text)) return '台湾';
  if (/^(?:kr|韓国|韩国|韓語|韩语)$/i.test(text) || /韓国|韩国|韓語|韩语/.test(text)) return '韓国';
  if (/^(?:jp|日本|日本語)$/i.test(text) || /日本/.test(text)) return '日本';
  if (/^(?:cn|中国|中國|簡體|简体)$/i.test(text) || /中国|中國|簡體|简体/.test(text)) return '中国';
  if (/^(?:en|英語|英语)$/i.test(text) || /英語|英语|english/i.test(text)) return '英語';
  return text;
}

function normalizeDictionaryCategory_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  // まんが: CM=コミック, CB=コミックブック
  if (/^(?:CM|CB)$/i.test(text)) return 'まんが';
  if (/まんが|マンガ|漫画|漫畫|comic|manga|webtoon|コミック|만화/i.test(text)) return 'まんが';
  // グッズ: GD=グッズ, STK=ステッカー
  if (/^(?:GD|STK)$/i.test(text)) return 'グッズ';
  if (/goods|グッズ|굿즈/i.test(text)) return 'グッズ';
  // 書籍: NV=ノベル, ES=エッセイ, ART=アートブック, SC=設定集, DAIHON=台本, KIRIE=切り絵
  if (/^(?:NV|ES|ART|SC|DAIHON|KIRIE|DVD|OSTD)$/i.test(text)) return '書籍';
  if (/書籍|本|小説|小說|小说|設定集|アートブック|絵本|台本|エッセイ|シナリオ|book|novel/i.test(text)) return '書籍';
  // 雑誌: MZ
  if (/^MZ$/i.test(text)) return '雑誌';
  if (/雑誌|magazine|잡지/i.test(text)) return '雑誌';
  return '';
}
function normalizeDictionaryTitleKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000"'“”‘’・･:：!！?？,，、.．\-ー‐―–—~〜/／\\|()\[\]{}【】「」『』<>]/g, '');
}

function normalizeDictionaryAuthorKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000・･·.,，、]/g, '');
}

function parseDictionaryAliases_(value) {
  return String(value || '')
    .split(/[\r\n;；,，、/／|｜]+/)
    .map(function(part) { return String(part || '').trim(); })
    .filter(Boolean);
}

function normalizeGoodsDictionaryTitle_(value) {
  const fallback = String(value || '').trim().replace(/^[台臺][湾灣]版\s*/u, '');
  if (!fallback) return '';

  const mediaTagged = fallback.match(/^(.+?)\s*[（(][^）)]*(?:漫畫|漫画|コミック|小說|小说|小説|動畫|动画|アニメ|影集|ドラマ|電影|映画|劇場版)[^）)]*[)）](?:\s+.*)?$/u);
  if (mediaTagged && mediaTagged[1]) return String(mediaTagged[1]).trim();

  const merchTagged = fallback.match(/^(.+?)(?:\s+|[-－–—:：・‧·|｜/／]|[A-Za-z0-9]-)+[^・‧·|｜/／\s:：\-－–—]*?(?:全棉托特袋|托特袋|帆布袋|手提袋|多功能手機包|手機包|手機袋|手機殼|手機壳|手機架|手机包|手机袋|手机壳|手机架|透明立方|壓克力(?:牌|立牌|磚|砖|便利夾|筋牌)?|压克力(?:牌|立牌|磚|砖|便利夹|筋牌)?|便利夾|便利夹|立牌|海報|海报|掛軸|挂轴|徽章|胸章|卡片|貼紙|贴纸|明信片|抱枕|滑鼠墊|鼠標墊|鼠标垫|マウスパッド|桌墊|桌垫|吊飾|吊饰|鑰匙圈|钥匙圈|杯墊|杯垫|色紙|色纸|套組|套装|福袋|拼圖|拼图|公仔|玩偶|資料夾|资料夹|文件夾|文件夹|T恤|毛巾|桌曆|桌历|年曆|年历)(?:\s*[\d０-９]{1,4}\s*入)?/u);
  if (merchTagged && merchTagged[1]) return String(merchTagged[1]).trim();

  return fallback;
}

function extractGoodsDictionaryOriginalTitle_(rowData) {
  const directCandidates = [
    rowData['作品名（原題）'],
    rowData['作品名原題'],
    rowData['原題タイトル'],
  ];

  for (let i = 0; i < directCandidates.length; i += 1) {
    const candidate = String(directCandidates[i] || '').trim();
    if (candidate) return candidate;
  }

  const rawTitle = String(
    rowData['商品名（原題）'] ||
    rowData['原題商品タイトル'] ||
    rowData['商品名'] ||
    rowData['タイトル'] ||
    ''
  ).trim();
  return normalizeGoodsDictionaryTitle_(rawTitle);
}

function resolveJapaneseTitleLookupFields_(item) {
  const rowData = extractRowData_(item);
  const itemType = String(item && item.itemType || '').trim();

  if (!rowData) return [];

  const language = normalizeDictionaryLanguage_(rowData['言語'] || '');
  const author = String(rowData['著者'] || rowData['作者'] || '').trim();
  const originalProductTitle = String(
    rowData['原題商品タイトル'] ||
    rowData['商品名（原題）'] ||
    rowData['商品名'] ||
    rowData['タイトル'] ||
    rowData['原題タイトル'] ||
    ''
  ).trim();
  if (!language) return [];

  if (itemType === 'goods') {
    const originalTitle = extractGoodsDictionaryOriginalTitle_(rowData);
    if (!originalTitle) return [];
    return [{
      language: language,
      category: 'まんが',
      originalTitle: originalTitle,
      author: author,
      originalProductTitle: originalProductTitle,
    }, {
      language: language,
      category: 'グッズ',
      originalTitle: originalTitle,
      author: author,
      originalProductTitle: originalProductTitle,
    }];
  }

  const category = normalizeDictionaryCategory_(rowData['カテゴリ'] || '') ||
    (itemType === 'magazine' ? '雑誌' : '書籍');

  const originalTitle = String(rowData['原題タイトル'] || rowData['原題商品タイトル'] || rowData['タイトル'] || '').trim();
  if (!originalTitle) return [];

  return [{
    language: language,
    category: category,
    originalTitle: originalTitle,
    author: author,
    originalProductTitle: originalProductTitle,
  }];
}

function findTitleFromDictionary_(language, category, originalTitle, author, originalProductTitle) {
  const normalizedLanguage = normalizeDictionaryLanguage_(language);
  const normalizedCategory = normalizeDictionaryCategory_(category);
  const titleKeys = buildDictionaryTitleKeys_(originalTitle, normalizedLanguage, normalizedCategory, originalProductTitle);
  if (!normalizedLanguage || !normalizedCategory || !titleKeys.length) return '';

  for (let i = 0; i < titleKeys.length; i += 1) {
    const builtin = BUILTIN_TITLE_ALIAS_DICTIONARY[[normalizedLanguage, normalizedCategory, titleKeys[i]].join('|')];
    if (builtin) return builtin;
  }

  const dictionary = getJapaneseTitleDictionaryIndex_();
  return findTitleInDictionaryByKeyMap_(dictionary.byKey, language, category, originalTitle, author, originalProductTitle);
}

function isJapaneseTitleLookupStatusValue_(value) {
  const text = String(value || '').trim();
  if (!text) return false;
  if (text === JAPANESE_TITLE_NOT_LOOKED_UP_LABEL) return true;
  if (text === JAPANESE_TITLE_NOT_FOUND_LABEL) return true;
  if (text === JAPANESE_TITLE_LOOKUP_FAILED_LABEL) return true;
  if (/^登録なし(?:\(|$)/.test(text)) return true;
  if (/^照会失敗(?:\(|$)/.test(text)) return true;
  return false;
}

function buildJapaneseTitleLookupStatusLabel_(lookupResult) {
  const result = lookupResult || {};
  const failedSources = uniqNonEmptyTitles_(Array.isArray(result.failedSources) ? result.failedSources : []);
  if (failedSources.length) {
    return JAPANESE_TITLE_LOOKUP_FAILED_LABEL + '(' + failedSources.join('/') + ')';
  }

  const checkedSources = uniqNonEmptyTitles_(Array.isArray(result.checkedSources) ? result.checkedSources : []);
  if (checkedSources.length) {
    return JAPANESE_TITLE_NOT_FOUND_ALL_SOURCES_LABEL;
  }

  return JAPANESE_TITLE_NOT_FOUND_LABEL;
}

function inferLookupFailureSource_(error) {
  const explicitSource = String((error && error.lookupSource) || '').trim();
  if (explicitSource) return explicitSource;

  const message = String((error && (error.message || error.toString())) || '');
  if (!message) return '不明';
  let match = message.match(/MangaUpdates login error:\s*(\d+)/i);
  if (match) return 'MU login ' + match[1];
  if (/MangaUpdates login error:\s*session token missing/i.test(message)) return 'MU session token missing';
  match = message.match(/MangaUpdates API (search|detail) error:\s*(\d+)/i);
  if (match) return 'MU API ' + match[1].toLowerCase() + ' ' + match[2];
  match = message.match(/MangaUpdates API error:\s*(\d+)/i);
  if (match) return 'MU API ' + match[1];
  match = message.match(/MangaUpdates site search error:\s*(\d+)/i);
  if (match) return 'MU site ' + match[1];
  if (/MangaUpdates parse error:\s*login/i.test(message)) return 'MU parse login';
  if (/MangaUpdates parse error:\s*search/i.test(message)) return 'MU parse search';
  if (/MangaUpdates parse error:\s*detail/i.test(message)) return 'MU parse detail';
  if (/MangaUpdates parse error:/i.test(message)) return 'MU parse';
  if (/MangaUpdates|\bMU\b/i.test(message)) return 'MU';
  if (/AniList/i.test(message)) return 'AniList';
  if (/BookWalker/i.test(message)) return 'BookWalker';
  if (/Amazon/i.test(message)) return 'Amazon';
  if (/BL補助|BL\s*sites|DLsite|Renta|ちるちる/i.test(message)) return 'BL補助';
  if (/辞書|dictionary/i.test(message)) return '辞書';
  if (/lookup\s*fields|前処理|extract|normalize/i.test(message)) return '前処理';
  return '不明';
}

function buildJapaneseTitleLookupFailureLabel_(lookupResult, error) {
  const result = lookupResult || {};
  const aggregateResult = {
    checkedSources: uniqNonEmptyTitles_(Array.isArray(result.checkedSources) ? result.checkedSources : []),
    failedSources: uniqNonEmptyTitles_(Array.isArray(result.failedSources) ? result.failedSources : []),
  };
  const inferredSource = inferLookupFailureSource_(error);
  if (inferredSource) aggregateResult.failedSources.push(inferredSource);

  const statusLabel = buildJapaneseTitleLookupStatusLabel_(aggregateResult);
  if (statusLabel && statusLabel !== JAPANESE_TITLE_NOT_FOUND_LABEL) {
    return statusLabel;
  }
  return JAPANESE_TITLE_LOOKUP_FAILED_LABEL + '(' + (inferredSource || '不明') + ')';
}

function convertJapaneseTitle_(item) {
  if (!item) return item;

  try {
    const originalRowData = extractRowData_(item);
    const existingJapaneseTitle = String(originalRowData['日本語タイトル'] || '').trim();
    if (existingJapaneseTitle && !isJapaneseTitleLookupStatusValue_(existingJapaneseTitle)) {
      const rowData = finalizeSheetWorkTitleColumns_(Object.assign({}, originalRowData));
      return Object.assign({}, item, { rowData: rowData });
    }

    const lookupFieldsList = resolveJapaneseTitleLookupFields_(item);
    if (!lookupFieldsList.length) {
      const rowData = Object.assign({}, originalRowData, {
        '日本語タイトル': JAPANESE_TITLE_NOT_LOOKED_UP_LABEL,
      });
      return Object.assign({}, item, { rowData: rowData });
    }

    let japaneseTitle = '';
    let lookupError = null;
    const aggregateLookupResult = {
      checkedSources: [],
      failedSources: [],
    };
    for (let i = 0; i < lookupFieldsList.length; i += 1) {
      const fields = lookupFieldsList[i];
      try {
        const lookupResult = cascadeLookupJapaneseTitleResult_(fields.language, fields.category, fields.originalTitle, fields.author, fields.originalProductTitle);
        aggregateLookupResult.checkedSources = aggregateLookupResult.checkedSources.concat(lookupResult.checkedSources || []);
        aggregateLookupResult.failedSources = aggregateLookupResult.failedSources.concat(lookupResult.failedSources || []);
        japaneseTitle = lookupResult.japaneseTitle || '';
      } catch (error) {
        lookupError = error;
        const failedSource = inferLookupFailureSource_(error);
        if (failedSource) {
          aggregateLookupResult.failedSources.push(failedSource);
        }
        Logger.log('Japanese title lookup failed: ' + (error && error.message ? error.message : error));
      }
      if (japaneseTitle) break;
    }

    const statusLabel = japaneseTitle
      ? ''
      : (lookupError
        ? buildJapaneseTitleLookupFailureLabel_(aggregateLookupResult, lookupError)
        : buildJapaneseTitleLookupStatusLabel_(aggregateLookupResult));

    const rowData = Object.assign({}, originalRowData, {
      '日本語タイトル': japaneseTitle || statusLabel,
    });

    if (japaneseTitle && String(item.itemType || '') === 'goods') {
      if (!String(rowData['作品名（日本語）'] || '').trim()) {
        rowData['作品名（日本語）'] = japaneseTitle;
      }
      if (!String(rowData['作品名日本語'] || '').trim()) {
        rowData['作品名日本語'] = japaneseTitle;
      }
    }

    return Object.assign({}, item, { rowData: finalizeSheetWorkTitleColumns_(rowData) });
  } catch (error) {
    Logger.log('Japanese title dictionary conversion failed: ' + (error && error.message ? error.message : error));
    const fallbackRowData = Object.assign({}, extractRowData_(item), {
      '日本語タイトル': buildJapaneseTitleLookupFailureLabel_({ checkedSources: [], failedSources: [] }, error),
    });
    return Object.assign({}, item, { rowData: fallbackRowData });
  }
}

// ==========================================
// 日本語タイトル辞書検索（初期実装）
// まんが / グッズのみ辞書変換
// 見つからなければ元の値をそのまま使う
// ==========================================
// カスケード日本語タイトル検索
// 辞書 → MangaUpdates → AniList → BookWalker → (BLのみ) BL補助
// 強い候補が見つかった時点で即終了
// ==========================================

function normalizeProviderId_(providerId) {
  const text = String(providerId || '').trim();
  if (!text) return '';
  if (TITLE_LOOKUP_PROVIDERS[text]) return text;
  const key = text.replace(/[\s_\-]/g, '').toLowerCase();
  return PROVIDER_ID_ALIASES[key] || text;
}

function normalizeProviderOrder_(providerOrder) {
  return uniqNonEmptyTitles_((providerOrder || []).map(function(providerId) {
    return normalizeProviderId_(providerId);
  })).filter(function(providerId) {
    return !!providerId;
  });
}

function inferLookupItemType_(itemType, category, originalTitle, author) {
  const explicit = String(itemType || '').trim();
  if (explicit) return explicit;
  const normalizedCategory = normalizeDictionaryCategory_(category);
  if (isBLLike_(originalTitle, category, author)) return 'bl_manga';
  if (normalizedCategory === 'グッズ') return 'goods';
  if (normalizedCategory === '雑誌') return 'magazine';
  if (normalizedCategory === '書籍') return 'novel_book';
  if (normalizedCategory === 'まんが') return 'manga';
  return 'unknown';
}

function providerOrderForLookup_(itemType, category, originalTitle, author, providerHint) {
  const inferredItemType = inferLookupItemType_(itemType, category, originalTitle, author);
  const hinted = providerHint && Array.isArray(providerHint.preferredOrder)
    ? normalizeProviderOrder_(providerHint.preferredOrder)
    : [];
  if (hinted.length) return hinted;
  return PROVIDER_ORDER_BY_ITEM_TYPE[inferredItemType] || PROVIDER_ORDER_BY_ITEM_TYPE.unknown;
}

function isProviderApplicableForItemType_(providerId, itemType) {
  if (providerId === 'workTitleMaster' || providerId === 'titleAliasDictionary') return true;
  const provider = TITLE_LOOKUP_PROVIDERS[providerId];
  if (!provider || !Array.isArray(provider.itemTypes) || !provider.itemTypes.length) return true;
  return provider.itemTypes.indexOf(itemType) >= 0;
}

function runTitleLookupProvider_(providerId, context) {
  switch (providerId) {
    case 'workTitleMaster': {
      const index = getWorkTitleMasterIndex_();
      if (!index.available) {
        return { japaneseTitle: '', candidates: [], skippedReason: 'sheet_missing' };
      }
      const japaneseTitle = findTitleInDictionaryByKeyMap_(
        index.byKey,
        context.language,
        context.category,
        context.originalTitle,
        context.author,
        context.originalProductTitle
      );
      return { japaneseTitle: japaneseTitle, candidates: [] };
    }
    case 'titleAliasDictionary': {
      const japaneseTitle = findTitleFromDictionary_(
        context.language,
        context.category,
        context.originalTitle,
        context.author,
        context.originalProductTitle
      );
      return { japaneseTitle: japaneseTitle, candidates: [] };
    }
    case 'chilchil': {
      if (!isBLLike_(context.originalTitle, context.category, context.author)) {
        return { japaneseTitle: '', candidates: [], skippedReason: 'not_applicable' };
      }
      const japaneseTitle = searchBLSites_(context.originalTitle, context.candidates || []);
      return { japaneseTitle: japaneseTitle || '', candidates: [] };
    }
    case 'mangaUpdates':
      return searchMangaUpdatesCandidates_(
        context.originalTitle,
        context.language,
        context.category,
        context.originalProductTitle
      );
    case 'aniList':
      return searchAniListCandidates_(
        context.originalTitle,
        context.language,
        context.category,
        context.originalProductTitle
      );
    case 'bookWalker': {
      const japaneseTitle = confirmOnBookWalker_(
        context.originalTitle,
        context.candidates || [],
        context.language,
        context.category,
        context.originalProductTitle
      );
      return { japaneseTitle: japaneseTitle || '', candidates: [] };
    }
    case 'amazon': {
      const japaneseTitle = confirmOnAmazonJp_(
        context.originalTitle,
        context.candidates || [],
        context.language,
        context.category,
        context.originalProductTitle
      );
      return { japaneseTitle: japaneseTitle || '', candidates: [] };
    }
    default:
      return { japaneseTitle: '', candidates: [], skippedReason: 'not_implemented' };
  }
}

function cascadeLookupJapaneseTitleResult_(language, category, originalTitle, author, originalProductTitle, itemType, providerHint, skipExternalApi) {
  const result = {
    japaneseTitle: '',
    source: '',
    checkedSources: [],
    failedSources: [],
    candidates: [],
    skippedSources: [],
    trace: [],
    errors: [],
  };
  if (!originalTitle) return result;

  const lookupItemType = inferLookupItemType_(itemType, category, originalTitle, author);
  const providerOrder = providerOrderForLookup_(lookupItemType, category, originalTitle, author, providerHint);
  const context = {
    language: language,
    category: category,
    originalTitle: originalTitle,
    author: author,
    originalProductTitle: originalProductTitle,
    itemType: lookupItemType,
    candidates: []
  };

  for (let i = 0; i < providerOrder.length; i += 1) {
    const providerId = normalizeProviderId_(providerOrder[i]);
    const provider = TITLE_LOOKUP_PROVIDERS[providerId];
    if (!provider) {
      result.skippedSources.push(providerId || 'unknown');
      result.trace.push((providerId || 'unknown') + ':skipped(unknown_provider)');
      continue;
    }
    if (!provider.implemented) {
      result.skippedSources.push(providerId);
      result.trace.push(providerId + ':skipped(not_implemented)');
      continue;
    }
    if (provider.delegatedToExtension) {
      result.skippedSources.push(providerId);
      result.trace.push(providerId + ':delegated_to_extension');
      continue;
    }
    if (!isProviderApplicableForItemType_(providerId, lookupItemType)) {
      result.skippedSources.push(providerId);
      result.trace.push(providerId + ':skipped(not_applicable)');
      continue;
    }
    if (skipExternalApi && providerId !== 'workTitleMaster' && providerId !== 'titleAliasDictionary') {
      result.skippedSources.push(providerId);
      result.trace.push(providerId + ':skipped(skip_external_during_write)');
      continue;
    }

    result.checkedSources.push(providerId);
    try {
      const providerResult = runTitleLookupProvider_(providerId, context) || {};
      if (providerResult.skippedReason) {
        result.skippedSources.push(providerId);
        result.trace.push(providerId + ':skipped(' + providerResult.skippedReason + ')');
        continue;
      }

      const providerCandidates = uniqNonEmptyTitles_(providerResult.candidates || []);
      if (providerCandidates.length) {
        result.candidates = uniqNonEmptyTitles_(result.candidates.concat(providerCandidates));
        context.candidates = uniqNonEmptyTitles_(context.candidates.concat(providerCandidates));
      }

      if (providerResult.japaneseTitle) {
        result.japaneseTitle = providerResult.japaneseTitle;
        result.source = providerId;
        result.trace.push(providerId + ':hit');
        return result;
      }

      result.trace.push(providerCandidates.length
        ? providerId + ':miss(candidates=' + providerCandidates.slice(0, 3).join('/') + ')'
        : providerId + ':miss');
    } catch (e) {
      const message = e && e.message ? e.message : String(e || 'unknown error');
      const failedSource = inferLookupFailureSource_(e) || providerId;
      result.failedSources.push(providerId);
      result.errors.push(providerId + ':' + message);
      result.trace.push(providerId + ':failed(' + failedSource + ')');
      Logger.log(providerId + ' cascade error: ' + message);
      if (!provider.failSoft) continue;
    }
  }

  result.candidates = uniqNonEmptyTitles_(result.candidates);
  const bestCandidate = pickBestCandidate_(result.candidates) || '';
  if (bestCandidate) {
    result.japaneseTitle = bestCandidate;
    result.source = '候補';
    result.trace.push('candidate:hit');
  }

  return result;
}

function cascadeLookupJapaneseTitle_(language, category, originalTitle, author, originalProductTitle) {
  return cascadeLookupJapaneseTitleResult_(language, category, originalTitle, author, originalProductTitle).japaneseTitle || '';
}

function isBLLike_(originalTitle, category, author) {
  const text = [originalTitle || '', category || '', author || ''].join(' ').toLowerCase();
  if (/bl|ボーイズラブ|耽美|danmei|yaoi|boys'?\s*love/.test(text)) return true;
  return false;
}


// ==========================================
// ② MangaUpdates 候補収集
// ==========================================

function searchMangaUpdatesCandidates_(query, language, category, originalProductTitle) {
  const result = { japaneseTitle: '', candidates: [] };
  const cacheKey = normalizeMangaUpdatesTitleKey_([
    normalizeDictionaryLanguage_(language),
    normalizeDictionaryCategory_(category),
    originalProductTitle || '',
    query,
  ].join('::'));
  if (!cacheKey) return result;

  const cache = CacheService.getScriptCache();
  const cached = cache.get('mu_' + cacheKey);
  if (cached !== null) {
    if (cached === MANGA_UPDATES_EMPTY_CACHE_VALUE) return result;
    result.japaneseTitle = cached;
    return result;
  }

  const searchQueries = buildMangaUpdatesSearchQueries_(query, language, category, originalProductTitle).slice(0, 8);
  const seenSeriesIds = {};
  let lastApiError = null;

  for (let q = 0; q < searchQueries.length; q += 1) {
    let searchResponse = null;
    try {
      searchResponse = fetchMangaUpdatesJson_('/series/search', {
        method: 'post',
        payload: {
          search: searchQueries[q],
          stype: 'title',
          page: 1,
          perpage: 5,
        },
      });
    } catch (error) {
      lastApiError = error;
      continue;
    }
    const results = Array.isArray(searchResponse && searchResponse.results)
      ? searchResponse.results.slice(0, 5)
      : [];

    for (let i = 0; i < results.length; i += 1) {
      const record = results[i];
      const seriesId = record && record.record ? record.record.series_id : '';
      if (!seriesId || seenSeriesIds[seriesId]) continue;
      seenSeriesIds[seriesId] = true;

      let detail = null;
      try {
        detail = fetchMangaUpdatesJson_('/series/' + seriesId, { method: 'get' });
      } catch (error) {
        lastApiError = error;
        continue;
      }
      if (!isStrongMangaUpdatesMatch_(query, record, detail, language, category, originalProductTitle)) continue;

      const associated = Array.isArray(detail && detail.associated) ? detail.associated : [];
      associated.forEach(function(item) {
        if (item && item.title) result.candidates.push(String(item.title).trim());
      });
      if (detail && detail.title) result.candidates.push(String(detail.title).trim());
      if (record && record.record && record.record.title) {
        result.candidates.push(String(record.record.title).trim());
      }

      const jpTitle = pickJapaneseMangaUpdatesTitle_(record, detail);
      if (jpTitle) {
        result.japaneseTitle = jpTitle;
        cache.put('mu_' + cacheKey, jpTitle, MANGA_UPDATES_CACHE_TTL_SECONDS);
        return result;
      }
    }
  }

  const siteFallback = searchMangaUpdatesSiteCandidates_(
    searchQueries,
    query,
    language,
    category,
    originalProductTitle
  );
  if (siteFallback.candidates && siteFallback.candidates.length) {
    result.candidates = result.candidates.concat(siteFallback.candidates);
  }
  if (!result.japaneseTitle && siteFallback.japaneseTitle) {
    result.japaneseTitle = siteFallback.japaneseTitle;
    cache.put('mu_' + cacheKey, siteFallback.japaneseTitle, MANGA_UPDATES_CACHE_TTL_SECONDS);
    result.candidates = uniqNonEmptyTitles_(result.candidates);
    return result;
  }
  result.candidates = uniqNonEmptyTitles_(result.candidates);
  if (!result.japaneseTitle && lastApiError && !result.candidates.length) {
    throw lastApiError;
  }
  if (!result.japaneseTitle && siteFallback.failed && siteFallback.error && !result.candidates.length
      && !/\b403\b/.test(siteFallback.error)) {
    throw new Error(siteFallback.error);
  }
  if (!result.japaneseTitle) {
    cache.put('mu_' + cacheKey, MANGA_UPDATES_EMPTY_CACHE_VALUE, MANGA_UPDATES_CACHE_TTL_SECONDS);
  }
  return result;
}

function searchMangaUpdatesSiteCandidates_(queries, query, language, category, originalProductTitle) {
  const result = { japaneseTitle: '', candidates: [], failed: false, error: '' };
  const searchQueries = uniqNonEmptyTitles_(queries || []).slice(0, 4);
  if (!searchQueries.length) return result;

  const queryKeys = buildMangaUpdatesTitleKeys_(query, language, category, originalProductTitle);
  let hadResponse = false;
  let lastError = '';

  // fetchAll で全クエリを並列実行
  const requests = searchQueries.map(function(sq) {
    return {
      url: MANGA_UPDATES_SITE_SEARCH_BASE + encodeURIComponent(sq),
      method: 'GET',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        Accept: 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.9,en-US;q=0.8,zh-TW;q=0.6',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        Origin: 'https://www.mangaupdates.com',
        Referer: 'https://www.mangaupdates.com/',
      },
    };
  });

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
    const detail = error && error.message ? error.message : String(error || 'unknown error');
    result.failed = true;
    result.error = 'MangaUpdates site search error: ' + detail;
    return result;
  }

  for (let i = 0; i < responses.length; i += 1) {
    try {
      const status = responses[i].getResponseCode();
      if (status < 200 || status >= 300) {
        lastError = 'MangaUpdates site search error: ' + status;
        continue;
      }

      hadResponse = true;
      const titles = extractMangaUpdatesSiteSeriesTitles_(responses[i].getContentText() || '');
      if (!titles.length) continue;

      result.candidates = result.candidates.concat(titles);
      if (!result.japaneseTitle) {
        const japaneseTitle = pickMatchingMangaUpdatesSiteJapaneseTitle_(
          titles, queryKeys, language, category, originalProductTitle
        );
        if (japaneseTitle) {
          result.japaneseTitle = japaneseTitle;
        }
      }
    } catch (error) {
      const detail = error && error.message ? error.message : String(error || 'unknown error');
      lastError = /^MangaUpdates/i.test(detail)
        ? detail
        : 'MangaUpdates site search error: ' + detail;
    }
  }

  result.candidates = uniqNonEmptyTitles_(result.candidates);
  if (!hadResponse && lastError) {
    result.failed = true;
    result.error = lastError;
  }
  return result;
}

function extractMangaUpdatesSiteSeriesTitles_(html) {
  const raw = String(html || '');
  const titles = [];
  const seen = {};

  function pushTitle_(fragment) {
    const text = normalizeTitleSearchCandidateSpacing_(
      decodeHtmlEntities_(String(fragment || '').replace(/<[^>]+>/g, ' '))
    );
    if (!text || text.length < 2 || seen[text]) return;
    seen[text] = true;
    titles.push(text);
  }

  const pattern = /<a\b[^>]*href="[^"]*(?:\/series\/|series\.html\?id=)[^"]*"[^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = pattern.exec(raw)) !== null) {
    pushTitle_(match[1]);
  }

  const underlineRe = /class="[^"]*linked-name-module__[^"]*__name_underline[^"]*"[^>]*>([^<]{1,240})</gi;
  while ((match = underlineRe.exec(raw)) !== null) {
    pushTitle_(match[1]);
  }

  return titles;
}

function pickMatchingMangaUpdatesSiteJapaneseTitle_(titles, queryKeys, language, category, originalProductTitle) {
  const uniqueTitles = uniqNonEmptyTitles_(titles || []);
  const japaneseTitles = uniqueTitles.filter(function(title) {
    return hasJapaneseTitleSignal_(title);
  });

  for (let i = 0; i < japaneseTitles.length; i += 1) {
    const candidateKeys = buildMangaUpdatesTitleKeys_(japaneseTitles[i], language, category, originalProductTitle);
    const matched = candidateKeys.some(function(candidateKey) {
      return (queryKeys || []).some(function(queryKey) {
        if (!candidateKey || !queryKey) return false;
        if (candidateKey === queryKey) return true;
        if (candidateKey.length >= 4 && queryKey.length >= 4) {
          return candidateKey.indexOf(queryKey) !== -1 || queryKey.indexOf(candidateKey) !== -1;
        }
        return false;
      });
    });
    if (matched) return japaneseTitles[i];
  }

  return japaneseTitles.length ? japaneseTitles[0] : '';
}

function decodeHtmlEntities_(value) {
  return String(value || '')
    .replace(/&#x([0-9a-f]+);/gi, function(_, hex) { return String.fromCodePoint(parseInt(hex, 16)); })
    .replace(/&#(\d+);/g, function(_, dec) { return String.fromCodePoint(parseInt(dec, 10)); })
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}
function searchAniListCandidates_(query, language, category, originalProductTitle) {
  const result = { japaneseTitle: '', candidates: [] };
  if (!query) return result;

  // カテゴリに応じて検索タイプを決定（小説は MANGA と NOVEL 両方試す）
  const normalizedCategory = normalizeDictionaryCategory_(category);
  const mediaTypes = (normalizedCategory === '書籍' || /小説|小說|小说|novel|ノベル/i.test(category))
    ? ['MANGA', 'NOVEL']
    : ['MANGA'];

  // クエリ数を削減して高速化（8→4）
  const searchQueries = buildMangaUpdatesSearchQueries_(query, language, category, originalProductTitle).slice(0, 4);

  // fetchAll で全クエリ×メディアタイプを並列実行
  const requests = [];
  const requestMeta = [];
  for (let q = 0; q < searchQueries.length; q += 1) {
    for (let t = 0; t < mediaTypes.length; t += 1) {
      const graphqlQuery = '{ Page(page: 1, perPage: 5) { media(search: ' + JSON.stringify(searchQueries[q])
        + ', type: ' + mediaTypes[t] + ') { '
        + 'title { romaji english native userPreferred } '
        + 'synonyms '
        + '} } }';
      requests.push({
        url: ANILIST_API_URL,
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ query: graphqlQuery }),
        muteHttpExceptions: true,
      });
      requestMeta.push({ queryIndex: q, mediaType: mediaTypes[t] });
    }
  }

  if (!requests.length) return result;

  const responses = UrlFetchApp.fetchAll(requests);
  for (let r = 0; r < responses.length; r += 1) {
    const response = responses[r];
    if (response.getResponseCode() !== 200) continue;

    const data = JSON.parse(response.getContentText() || '{}');
    const page = data && data.data ? data.data.Page : null;
    const mediaList = page && Array.isArray(page.media) ? page.media : [];
    if (!mediaList.length) continue;

    for (let mi = 0; mi < mediaList.length; mi += 1) {
      const media = mediaList[mi];
      if (!media) continue;

      const title = media.title || {};
      const titleValues = [title.native, title.romaji, title.english, title.userPreferred];
      const synonyms = Array.isArray(media.synonyms) ? media.synonyms : [];
      const allTitles = uniqNonEmptyTitles_(titleValues.concat(synonyms));
      result.candidates = uniqNonEmptyTitles_(result.candidates.concat(allTitles));

      const nativeTitle = String(title.native || '').trim();
      if (nativeTitle && hasJapaneseTitleSignal_(nativeTitle)) {
        result.japaneseTitle = nativeTitle;
        return result;
      }

      for (let i = 0; i < allTitles.length; i += 1) {
        if (hasJapaneseTitleSignal_(allTitles[i])) {
          result.japaneseTitle = allTitles[i];
          return result;
        }
      }
    }
  }

  return result;
}

function isWesternLatinTitleQuery_(value) {
  const s = String(value || '').trim();
  if (s.length < 4 || s.length > 160) return false;
  if (!/[a-zA-Z]/.test(s)) return false;
  if (hasJapaneseTitleSignal_(s)) return false;
  if (/[\u4E00-\u9FFF\u3400-\u4DBF\u3040-\u30FF]/.test(s)) return false;
  if (/[가-힣]/.test(s)) return false;
  return true;
}

function confirmOnBookWalker_(originalQuery, candidates, language, category, originalProductTitle) {
  const japanCandidates = candidates.filter(function(c) { return hasJapaneseTitleSignal_(c); });
  const latinCandidates = uniqNonEmptyTitles_((candidates || []).filter(function(c) {
    return isWesternLatinTitleQuery_(c);
  }));
  const searchTerms = uniqNonEmptyTitles_(
    japanCandidates.concat(
      latinCandidates.slice(0, 4),
      buildMangaUpdatesSearchQueries_(originalQuery, language, category, originalProductTitle),
      [originalQuery]
    )
  ).slice(0, 8);

  // fetchAll で全検索語を並列実行
  const requests = searchTerms.map(function(term) {
    return {
      url: 'https://bookwalker.jp/search/?word=' + encodeURIComponent(term) + '&order=score',
      method: 'GET',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        Accept: 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.8',
        Referer: 'https://bookwalker.jp/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
      },
    };
  });
  if (!requests.length) return '';

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    Logger.log('BookWalker fetchAll error: ' + e.message);
    return '';
  }

  for (let i = 0; i < responses.length; i += 1) {
    try {
      if (responses[i].getResponseCode() !== 200) continue;
      const html = responses[i].getContentText() || '';
      const bwTitles = extractBookWalkerJapaneseTitlesFromHtml_(html).filter(function(t) {
        return hasJapaneseTitleSignal_(t);
      });
      for (let j = 0; j < Math.min(bwTitles.length, 12); j += 1) {
        const bwTitle = bwTitles[j];
        // 日本語候補が無い状態で検索結果の先頭を採用すると、
        // 無関係な作品（例: 週刊ダイヤモンド）を誤って resolved にしてしまう。
        if (!japanCandidates.length) continue;

        const bwKeys = buildMangaUpdatesTitleKeys_(bwTitle);
        for (let k = 0; k < japanCandidates.length; k += 1) {
          const candidate = japanCandidates[k];
          const candidateKeys = buildMangaUpdatesTitleKeys_(candidate);
          const matched = candidateKeys.some(function(candKey) {
            return bwKeys.some(function(bwKey) {
              if (!candKey || !bwKey) return false;
              return candKey === bwKey
                || (candKey.length >= 4 && bwKey.length >= 4 && (bwKey.indexOf(candKey) >= 0 || candKey.indexOf(bwKey) >= 0));
            });
          });
          if (matched) return candidate;
        }
      }
    } catch (e) {
      Logger.log('BookWalker search error: ' + e.message);
    }
  }
  return '';
}

function extractBookWalkerJapaneseTitlesFromHtml_(html) {
  const h = String(html || '');
  const titles = [];
  const seen = {};

  function addTitle_(value) {
    const t = String(value || '').replace(/\s+/g, ' ').trim();
    if (!t || t.length < 2 || seen[t]) return;
    seen[t] = true;
    titles.push(t);
  }

  const reTitleAttr = /<a\b[^>]*\bclass="[^"]*m-book-item__title[^"]*"[^>]*\btitle="([^"]+)"/gi;
  let m;
  while ((m = reTitleAttr.exec(h)) !== null) {
    addTitle_(m[1]);
  }

  const reInner = /<a\b[^>]*\bclass="[^"]*m-book-item__title[^"]*"[^>]*>([\s\S]*?)<\/a>/gi;
  while ((m = reInner.exec(h)) !== null) {
    addTitle_(String(m[1] || '').replace(/<[^>]+>/g, ' '));
  }

  const legacyMatches = h.match(/class="[^"]*item[Tt]itle[^"]*"[^>]*>([^<]+)</g) || [];
  for (let i = 0; i < legacyMatches.length; i += 1) {
    const inner = legacyMatches[i].match(/>([^<]+)$/);
    if (inner) addTitle_(inner[1]);
  }

  return titles;
}

// ==========================================
// ⑤ Amazon.co.jp 確認検索
// BookWalker と同じくMU/AniListの候補を確認する補助検索
// ==========================================

function confirmOnAmazonJp_(originalQuery, candidates, language, category, originalProductTitle) {
  const japanCandidates = candidates.filter(function(c) { return hasJapaneseTitleSignal_(c); });
  const searchTerms = uniqNonEmptyTitles_(
    japanCandidates.concat(
      buildMangaUpdatesSearchQueries_(originalQuery, language, category, originalProductTitle),
      [originalQuery]
    )
  ).slice(0, 3);

  // fetchAll で全検索語を並列実行
  const requests = searchTerms.map(function(term) {
    return {
      url: 'https://www.amazon.co.jp/s?k=' + encodeURIComponent(term) + '&i=stripbooks',
      method: 'GET',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'Accept-Language': 'ja',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
      },
    };
  });
  if (!requests.length) return '';

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    Logger.log('Amazon fetchAll error: ' + e.message);
    return '';
  }

  for (let i = 0; i < responses.length; i += 1) {
    try {
      if (responses[i].getResponseCode() !== 200) continue;
      const html = responses[i].getContentText() || '';
      let titleMatches = html.match(/class="[^"]*a-size-medium[^"]*"[^>]*>([^<]+)</g) || [];
      if (!titleMatches.length) {
        titleMatches = html.match(/class="[^"]*a-text-normal[^"]*"[^>]*>([^<]+)</g) || [];
      }

      for (let j = 0; j < Math.min(titleMatches.length, 5); j += 1) {
        const match = titleMatches[j].match(/>([^<]+)$/);
        if (!match) continue;
        const amzTitle = String(match[1]).trim();
        if (!hasJapaneseTitleSignal_(amzTitle)) continue;

        // 日本語候補が無い状態では、検索結果の先頭を確定タイトルとして採用しない。
        if (!japanCandidates.length) continue;

        const amzKeys = buildMangaUpdatesTitleKeys_(amzTitle);
        for (let k = 0; k < japanCandidates.length; k += 1) {
          const candidate = japanCandidates[k];
          const candidateKeys = buildMangaUpdatesTitleKeys_(candidate);
          const matched = candidateKeys.some(function(candKey) {
            return amzKeys.some(function(aKey) {
              if (!candKey || !aKey) return false;
              return candKey === aKey
                || (candKey.length >= 4 && aKey.length >= 4 && (aKey.indexOf(candKey) >= 0 || candKey.indexOf(aKey) >= 0));
            });
          });
          if (matched) return candidate;
        }
      }
    } catch (e) {
      Logger.log('Amazon search error: ' + e.message);
    }
  }
  return '';
}

function searchBLSites_(originalTitle, existingCandidates) {
  const japanCandidates = existingCandidates.filter(function(c) { return hasJapaneseTitleSignal_(c); });
  const searchTerms = uniqNonEmptyTitles_(japanCandidates.concat([originalTitle])).slice(0, 2);

  const blSites = [
    { name: 'DLsite', buildUrl: function(q) { return 'https://www.dlsite.com/bl-pro/fsr/=/keyword/' + encodeURIComponent(q); } },
    { name: 'Renta', buildUrl: function(q) { return 'https://renta.papy.co.jp/renta/sc/frm/search?word=' + encodeURIComponent(q); } },
    { name: 'Chil-chil', buildUrl: function(q) { return 'https://www.chil-chil.net/goodsList/find/?keyword=' + encodeURIComponent(q); } },
  ];

  // fetchAll で全サイト×全検索語を並列実行
  const requests = [];
  const requestMeta = [];
  for (let s = 0; s < blSites.length; s += 1) {
    for (let i = 0; i < searchTerms.length; i += 1) {
      requests.push({
        url: blSites[s].buildUrl(searchTerms[i]),
        method: 'GET',
        muteHttpExceptions: true,
        followRedirects: true,
      });
      requestMeta.push({ siteName: blSites[s].name });
    }
  }
  if (!requests.length) return '';

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    Logger.log('BL sites fetchAll error: ' + e.message);
    return '';
  }

  for (let r = 0; r < responses.length; r += 1) {
    try {
      if (responses[r].getResponseCode() !== 200) continue;
      const html = responses[r].getContentText() || '';
      const matches = html.match(/<(?:h[1-4]|a|span)[^>]*class="[^"]*(?:title|name|item)[^"]*"[^>]*>([^<]{2,80})</gi) || [];
      for (let j = 0; j < Math.min(matches.length, 5); j += 1) {
        const m = matches[j].match(/>([^<]+)$/);
        if (!m) continue;
        const blTitle = String(m[1]).trim();
        if (!hasJapaneseTitleSignal_(blTitle)) continue;

        const blKey = normalizeMangaUpdatesTitleKey_(blTitle);
        for (let k = 0; k < japanCandidates.length; k += 1) {
          const candKey = normalizeMangaUpdatesTitleKey_(japanCandidates[k]);
          if (candKey && blKey && (blKey.indexOf(candKey) >= 0 || candKey.indexOf(blKey) >= 0)) {
            return japanCandidates[k];
          }
        }
      }
    } catch (e) {
      Logger.log(requestMeta[r].siteName + ' search error: ' + e.message);
    }
  }
  return '';
}

// ==========================================
// 候補選択ユーティリティ
// ==========================================

function pickBestCandidate_(candidates) {
  if (!candidates || !candidates.length) return '';
  // 日本語シグナルがある候補を優先
  for (let i = 0; i < candidates.length; i += 1) {
    if (hasJapaneseTitleSignal_(candidates[i])) {
      return String(candidates[i]).trim();
    }
  }
  return '';
}
function testFindTitleFromDictionary_() {
  const result = findTitleFromDictionary_('台湾', 'まんが', '我內心的糟糕念頭', '');
  Logger.log(JSON.stringify({ japaneseTitle: result }, null, 2));
}

function testConvertJapaneseTitle_() {
  const item = {
    itemType: 'goods',
    rowData: {
      '言語': '台湾',
      '商品名（原題）': '鄰居是公會成員 壓克力牌 A',
      '原題商品タイトル': '鄰居是公會成員 壓克力牌 A',
      '日本語タイトル': '',
    },
  };
  Logger.log(JSON.stringify(convertJapaneseTitle_(item), null, 2));
}
const ANILIST_API_URL = 'https://graphql.anilist.co';
const MANGA_UPDATES_API_BASE = 'https://api.mangaupdates.com/v1';
const MANGA_UPDATES_SITE_SEARCH_BASE = 'https://www.mangaupdates.com/site/search/result?search=';
const MANGA_UPDATES_CACHE_TTL_SECONDS = 21600;
const MANGA_UPDATES_EMPTY_CACHE_VALUE = '__EMPTY__';
const MANGA_UPDATES_SESSION_CACHE_KEY = 'mu_session_token_v1';
const MANGA_UPDATES_SESSION_CACHE_TTL_SECONDS = 21600;
/** 推奨: スクリプトプロパティ MANGA_UPDATES_USERNAME / MANGA_UPDATES_PASSWORD を設定 */
const MANGA_UPDATES_USERNAME = 'niyantarose';
const MANGA_UPDATES_PASSWORD = 'niyantarose';

function getMangaUpdatesCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const username = String(props.getProperty('MANGA_UPDATES_USERNAME') || MANGA_UPDATES_USERNAME || '').trim();
  const password = String(props.getProperty('MANGA_UPDATES_PASSWORD') || MANGA_UPDATES_PASSWORD || '').trim();
  return { username: username, password: password };
}

/** 実行中に MU login が 403 で失敗したら以降のリクエストをスキップするフラグ（GoogleのIPブロック回避用） */
var __mangaUpdatesBlockedForRun = false;
function isMangaUpdatesBlockedForRun_() { return __mangaUpdatesBlockedForRun; }
function markMangaUpdatesBlockedForRun_() { __mangaUpdatesBlockedForRun = true; }


function handleMangaUpdatesLookupPost_(payload) {
  const queries = payload && Array.isArray(payload.queries) ? payload.queries : [];
  const source = payload && payload.source ? payload.source : 'unknown';
  if (!queries.length) {
    return jsonResponse_({
      ok: false,
      error: 'queries required',
      japaneseTitle: '',
      matchedQuery: '',
      via: 'mangaupdates_api',
      source: source,
    });
  }
  try {
    const resolved = lookupJapaneseTitleViaMangaUpdatesQueries_(queries);
    return jsonResponse_({
      ok: true,
      japaneseTitle: resolved.japaneseTitle || '',
      matchedQuery: resolved.matchedQuery || '',
      via: 'mangaupdates_api',
      source: source,
    });
  } catch (e) {
    const message = e && e.message ? e.message : String(e || 'unknown error');
    Logger.log('handleMangaUpdatesLookupPost_ error: ' + message);
    return jsonResponse_({
      ok: false,
      error: message,
      japaneseTitle: '',
      matchedQuery: '',
      via: 'mangaupdates_api',
      source: source,
    });
  }
}

function lookupJapaneseTitleViaMangaUpdatesQueries_(queries) {
  const normalizedQueries = uniqNonEmptyTitles_(queries || []);
  for (let i = 0; i < normalizedQueries.length; i += 1) {
    const query = normalizedQueries[i];
    const japaneseTitle = lookupJapaneseTitleViaMangaUpdates_(query);
    if (japaneseTitle) {
      return {
        japaneseTitle: japaneseTitle,
        matchedQuery: query,
      };
    }
  }

  return {
    japaneseTitle: '',
    matchedQuery: '',
  };
}

function lookupJapaneseTitleViaMangaUpdates_(query) {
  const cacheKey = normalizeMangaUpdatesTitleKey_(query);
  if (!cacheKey) return '';

  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached !== null) {
    return cached === MANGA_UPDATES_EMPTY_CACHE_VALUE ? '' : cached;
  }

  const searchResponse = fetchMangaUpdatesJson_('/series/search', {
    method: 'post',
    payload: {
      search: query,
      stype: 'title',
      page: 1,
      perpage: 5,
    },
  });
  const results = Array.isArray(searchResponse && searchResponse.results)
    ? searchResponse.results.slice(0, 3)
    : [];

  for (let i = 0; i < results.length; i += 1) {
    const result = results[i];
    const seriesId = result && result.record ? result.record.series_id : '';
    if (!seriesId) continue;

    const detail = fetchMangaUpdatesJson_('/series/' + seriesId, { method: 'get' });
    if (!isStrongMangaUpdatesMatch_(query, result, detail)) continue;

    const japaneseTitle = pickJapaneseMangaUpdatesTitle_(result, detail);
    if (japaneseTitle) {
      cache.put(cacheKey, japaneseTitle, MANGA_UPDATES_CACHE_TTL_SECONDS);
      return japaneseTitle;
    }
  }

  const siteFallback = searchMangaUpdatesSiteCandidates_([query], query, '', '', '');
  if (siteFallback.japaneseTitle) {
    cache.put(cacheKey, siteFallback.japaneseTitle, MANGA_UPDATES_CACHE_TTL_SECONDS);
    return siteFallback.japaneseTitle;
  }
  cache.put(cacheKey, MANGA_UPDATES_EMPTY_CACHE_VALUE, MANGA_UPDATES_CACHE_TTL_SECONDS);
  return '';
}

function fetchMangaUpdatesJson_(endpoint, options) {
  const endpointLabel = endpoint === '/series/search'
    ? 'search'
    : (/^\/series\/[^\/]+$/i.test(String(endpoint || ''))
        ? 'detail'
        : (String(endpoint || '').replace(/^\//, '') || 'unknown'));
  const method = String((options && options.method) || 'get').toUpperCase();
  const hasPayload = options && Object.prototype.hasOwnProperty.call(options, 'payload');
  const payloadObj = hasPayload ? (options.payload || {}) : null;
  const isLoginEndpoint = String(endpoint || '') === '/account/login';

  function buildRequestOpts(bearerToken) {
    const headers = { Accept: 'application/json' };
    if (bearerToken) {
      headers.Authorization = 'Bearer ' + bearerToken;
    }
    const ro = {
      method: method,
      muteHttpExceptions: true,
      headers: headers,
    };
    if (payloadObj !== null) {
      ro.contentType = 'application/json; charset=utf-8';
      ro.payload = JSON.stringify(payloadObj);
    }
    return ro;
  }

  function doFetch_(bearerToken) {
    return UrlFetchApp.fetch(MANGA_UPDATES_API_BASE + endpoint, buildRequestOpts(bearerToken));
  }

  if (!isLoginEndpoint && isMangaUpdatesBlockedForRun_()) {
    throw new Error('MangaUpdates API ' + endpointLabel + ' error: 403 (skipped: blocked for this run)');
  }

  var response;
  var authRetryNote = '';

  if (isLoginEndpoint) {
    try {
      response = doFetch_(null);
    } catch (e) {
      Logger.log('MangaUpdates API ' + endpointLabel + ' fetch error: ' + e.message);
      throw new Error('MangaUpdates API ' + endpointLabel + ' fetch error: ' + e.message);
    }
  } else {
    try {
      var bearer = getMangaUpdatesSessionToken_(false);
      response = doFetch_(bearer);
      authRetryNote = ' with bearer';
    } catch (e) {
      Logger.log('MangaUpdates API ' + endpointLabel + ' auth fetch error: ' + e.message);
      throw new Error('MangaUpdates API ' + endpointLabel + ' fetch error: ' + e.message);
    }
    var status = response.getResponseCode();
    if (status === 401 || status === 403) {
      try {
        clearMangaUpdatesSessionTokenCache_();
        var bearerFresh = getMangaUpdatesSessionToken_(true);
        response = doFetch_(bearerFresh);
        authRetryNote = ' after bearer refresh';
      } catch (authErr) {
        authRetryNote = ' after bearer refresh failed: ' + authErr.message;
        Logger.log('MangaUpdates API ' + endpointLabel + ' bearer refresh failed: ' + authErr.message);
      }
    }
  }

  var status = response.getResponseCode();
  var responseText = response.getContentText() || '';

  if (status === 403) {
    Logger.log('MangaUpdates API ' + endpointLabel + ' blocked (403)' + authRetryNote);
    throw new Error('MangaUpdates API ' + endpointLabel + ' error: 403' + authRetryNote);
  }
  if (status < 200 || status >= 300) {
    throw new Error('MangaUpdates API ' + endpointLabel + ' error: ' + status + authRetryNote);
  }

  try {
    return responseText ? JSON.parse(responseText) : {};
  } catch (error) {
    throw new Error('MangaUpdates parse error: ' + endpointLabel);
  }
}

function getMangaUpdatesSessionToken_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cachedToken = cache.get(MANGA_UPDATES_SESSION_CACHE_KEY);
    if (cachedToken) return cachedToken;
  }

  const creds = getMangaUpdatesCredentials_();
  if (!creds.username || !creds.password) {
    throw new Error('MangaUpdates credentials missing (set script properties or constants)');
  }

  const response = UrlFetchApp.fetch(MANGA_UPDATES_API_BASE + '/account/login', {
    method: 'put',
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8',
    headers: { Accept: 'application/json' },
    payload: JSON.stringify({
      username: creds.username,
      password: creds.password,
    }),
  });

  const status = response.getResponseCode();
  const responseText = response.getContentText() || '';
  if (status < 200 || status >= 300) {
    if (status === 403) markMangaUpdatesBlockedForRun_();
    throw new Error('MangaUpdates login error: ' + status);
  }

  let data = {};
  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch (error) {
    throw new Error('MangaUpdates parse error: login');
  }
  const sessionToken = String((((data || {}).context || {}).session_token) || '').trim();
  if (!sessionToken) {
    throw new Error('MangaUpdates login error: session token missing');
  }

  cache.put(MANGA_UPDATES_SESSION_CACHE_KEY, sessionToken, MANGA_UPDATES_SESSION_CACHE_TTL_SECONDS);
  return sessionToken;
}

function clearMangaUpdatesSessionTokenCache_() {
  CacheService.getScriptCache().remove(MANGA_UPDATES_SESSION_CACHE_KEY);
}

function buildMangaUpdatesSearchQueries_(value, language, category, originalProductTitle) {
  return buildTitleSearchCandidates_(language, category, value, originalProductTitle);
}

function buildTitleSearchCandidates_(language, category, originalTitle, originalProductTitle) {
  const normalizedLanguage = normalizeDictionaryLanguage_(language);
  const normalizedCategory = normalizeDictionaryCategory_(category);
  const queue = uniqNonEmptyTitles_([
    normalizeTitleSearchCandidateSpacing_(originalTitle),
    normalizeTitleSearchCandidateSpacing_(originalProductTitle),
  ]);
  const seen = {};
  const results = [];

  while (queue.length) {
    const current = normalizeTitleSearchCandidateSpacing_(queue.shift());
    if (!current) continue;

    const dedupeKey = normalizeDictionaryTitleKey_(current);
    if (!dedupeKey || seen[dedupeKey]) continue;
    seen[dedupeKey] = true;
    results.push(current);

    uniqNonEmptyTitles_([
      stripLanguageSpecificTitleNoise_(current, normalizedLanguage),
      stripCategorySpecificTitleNoise_(current, normalizedCategory),
      stripMangaUpdatesTitleNoise_(current, normalizedLanguage, normalizedCategory),
      stripLookupVolumeNoise_(current),
      normalizeLookupWorkTitle_(current),
      stripInlineEditionBrackets_(current),
      collapseSubtitleDelimiter_(current),
      stripMangaUpdatesEditionSuffix_(current),
      stripMangaUpdatesSubtitle_(current),
      stripMangaUpdatesSubtitle_(stripMangaUpdatesEditionSuffix_(current)),
      stripMangaUpdatesSubtitle_(collapseSubtitleDelimiter_(current)),
      stripMangaUpdatesEditionSuffix_(stripInlineEditionBrackets_(current)),
      stripMangaUpdatesSubtitle_(stripInlineEditionBrackets_(current)),
      stripMangaUpdatesEditionSuffix_(stripCategorySpecificTitleNoise_(current, normalizedCategory)),
      stripMangaUpdatesSubtitle_(stripCategorySpecificTitleNoise_(current, normalizedCategory)),
    ]).forEach(function(variant) {
      const cleaned = normalizeTitleSearchCandidateSpacing_(variant);
      const cleanedKey = normalizeDictionaryTitleKey_(cleaned);
      if (!cleaned || !cleanedKey || seen[cleanedKey]) return;
      queue.push(cleaned);
    });
  }

  return uniqNonEmptyTitles_(results).slice(0, 12);
}

function normalizeTitleSearchCandidateSpacing_(value) {
  return String(value || '')
    .replace(/[\u3000\t\r\n]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function stripLanguageSpecificTitleNoise_(value, language) {
  let text = normalizeTitleSearchCandidateSpacing_(value);
  if (!text) return '';

  const patterns = [];
  switch (normalizeDictionaryLanguage_(language)) {
    case '台湾':
    case '中国':
      patterns.push(
        /^(?:台灣版|臺灣版|台湾版|中文版|繁體中文版?|繁体中文版?|簡體中文版?|简体中文版?|中文書|中文书|韓版|韩版)\s*/u,
        /^(?:小說|小説|小说|漫畫|漫画|コミック|設定集|設定資料集|公式設定集|官方設定集|畫集|画集|藝術畫冊|艺术画册|アートブック|雜誌|杂志)\s*/u
      );
      break;
    case '日本':
      patterns.push(
        /^(?:日本版|日本語版|和書|邦訳版)\s*/u,
        /^(?:小説|ノベル|ライトノベル|ラノベ|漫画|コミック|設定集|設定資料集|画集|アートブック|雑誌)\s*/u
      );
      break;
    case '韓国':
      patterns.push(
        /^(?:한국판|한국어판|한글판)\s*/u,
        /^(?:소설|라이트노벨|만화|웹툰|설정집|화집|아트북|잡지)\s*/u
      );
      break;
    case '英語':
      patterns.push(
        /^(?:english edition|english ver\.?|eng(?:lish)? ver\.?|overseas edition)\s*/iu,
        /^(?:novel|light novel|manga|comic|art book|guide book|setting book|magazine)\s*/iu
      );
      break;
  }

  patterns.forEach(function(pattern) {
    text = text.replace(pattern, '').trim();
  });
  return text;
}

function stripCategorySpecificTitleNoise_(value, category) {
  let text = normalizeTitleSearchCandidateSpacing_(value);
  if (!text) return '';

  switch (normalizeDictionaryCategory_(category)) {
    case 'グッズ':
      text = text
        .replace(/\s*(?:全棉托特袋|托特袋|帆布袋|手提袋|壓克力(?:牌|立牌|磚|砖)?|压克力(?:牌|立牌|磚|砖)?|アクリル(?:スタンド|キーホルダー|プレート|ブロック)?|立牌|海報|海报|掛軸|挂轴|徽章|胸章|缶バッジ|卡片|貼紙|贴纸|ステッカー|シール|明信片|ポストカード|抱枕|滑鼠墊|鼠標墊|鼠标垫|マウスパッド|桌墊|桌垫|吊飾|吊饰|鑰匙圈|钥匙圈|キーホルダー|杯墊|杯垫|コースター|色紙|色纸|套組|套装|セット|福袋|拼圖|拼图|パズル|公仔|フィギュア|玩偶|ぬいぐるみ|資料夾|资料夹|文件夾|文件夹|クリアファイル|T恤|Ｔシャツ|毛巾|タオル|桌曆|桌历|年曆|年历|壓克力磁鐵|压克力磁铁|手機殼|手机壳)(?:\s+[A-Z0-9]+|\s*[\d０-９]{1,4}\s*入)?$/u, '')
        .trim();
      break;
    case '雑誌':
      text = text
        .replace(/\s*(?:\d{4}年\d{1,2}月號?|\d{4}年\d{1,2}月号|vol\.?\s*\d+|no\.?\s*\d+|第?\d+期)\s*$/iu, '')
        .trim();
      break;
  }

  return text;
}

function stripMangaUpdatesTitleNoise_(value, language, category) {
  return normalizeTitleSearchCandidateSpacing_(
    stripCategorySpecificTitleNoise_(
      stripLanguageSpecificTitleNoise_(String(value || '').trim(), language),
      category
    ).replace(/^[\[【(（][^\]】)）]{1,24}[\]】)）]\s*/u, '')
  );
}

function stripInlineEditionBrackets_(value) {
  return normalizeTitleSearchCandidateSpacing_(String(value || '')
    .replace(/\s*[（(【\[][^()（）\[\]【】]{0,24}(?:首刷|限定|版|特裝|特装|豪華|豪华|通常|套書|套装|附錄|附录|特典|初回|漫畫|漫画|コミック|小說|小说|小説|設定集|畫集|画集|アートブック)[^()（）\[\]【】]{0,24}[)）】\]]/gu, ' '));
}

function collapseSubtitleDelimiter_(value) {
  return normalizeTitleSearchCandidateSpacing_(String(value || '')
    .replace(/\s*[~〜～:：|｜／/]\s*/gu, ' '));
}

function stripMangaUpdatesSubtitle_(value) {
  const text = normalizeTitleSearchCandidateSpacing_(value);
  if (!text) return '';
  const match = text.match(/^(.+?)(?:\s*[~〜～:：|｜／/]\s*.+)$/u)
    || text.match(/^(.+?)(?:\s+[-–—]\s+.+)$/u);
  if (match && match[1] && match[1].trim().length >= 2) return match[1].trim();
  return text;
}

function stripMangaUpdatesEditionSuffix_(value) {
  return normalizeTitleSearchCandidateSpacing_(String(value || '')
    .replace(/\s*[（(]?(?:首刷限定版?|首刷限定|限定版?|特裝版?|特装版?|豪華版|豪华版|通常版|套書|套組|套装|附錄版|附录版|特典版|贈品版|初回限定?|台灣版|臺灣版|台湾版|韓版|韩版)[）)]?\s*$/u, '')
    .replace(/\s*(?:vol\.?\s*\d+|no\.?\s*\d+)\s*$/iu, '')
    .replace(/\s*[（(]?(?:第?\d+\s*(?:巻|卷|卷册|冊|册|集|期|話|号|號|漫畫|漫画|コミック|單行本|单行本)|\d+)[）)]?\s*$/u, ''));
}

function stripLookupWorkTitleDecoration_(value) {
  return normalizeTitleSearchCandidateSpacing_(String(value || '')
    .replace(/\s*[!！?？]{1,4}\s*$/u, ''));
}

function stripLookupVolumeNoise_(value) {
  var text = normalizeTitleSearchCandidateSpacing_(String(value || ''));
  if (!text) return '';
  var rules = [
    /\s*(?:vol\.?|v\.?|第)\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回|期|号|號)?\s*$/iu,
    /\s*[#＃]?\s*[0-9０-９]{1,4}\s*$/u,
    /\s*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/u,
    /\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/u,
    /\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/u
  ];
  for (var i = 0; i < 4; i += 1) {
    var prev = text;
    for (var r = 0; r < rules.length; r += 1) {
      text = text.replace(rules[r], '').trim();
    }
    if (text === prev) break;
  }
  return normalizeTitleSearchCandidateSpacing_(text);
}

/** シート出力·日本語タイトル列: (TV) 等のメディア表記を落とす（照会ステータス文字列は触れない） */
function sanitizeWorkNameJapaneseForSheet_(value) {
  var text = normalizeTitleSearchCandidateSpacing_(String(value || '').trim());
  if (!text) return '';
  if (isJapaneseTitleLookupStatusValue_(text)) return text;
  var i;
  for (i = 0; i < 8; i += 1) {
    var prev = text;
    text = text
      .replace(/\s*[\[(（〈【]\s*(?:TV|T\.V\.|OVA|OAD|ONA|SP|映画|電影|劇場版|剧场版|アニメ(?:版|ーション)?|ドラマ(?:版)?|真人版|Web\s*アニメ|Anime|Movie|Film)\s*[\])）〉】]/giu, '')
      .replace(/\s*[\[\(]\s*[Tt][Vv]\s*[\]\)]\s*$/u, '')
      .trim();
    text = normalizeTitleSearchCandidateSpacing_(text);
    if (text === prev) break;
  }
  return text;
}

/**
 * シート出力·原題タイトル列: 作品名のみ（巻・漫畫語尾・版括弧などを落とす）
 * ※lookup 用の normalizeLookupWorkTitle_ より控えめ（副題区切りで本編を切り捨てない）
 */
function sanitizeWorkNameOriginalForSheet_(value) {
  var text = normalizeTitleSearchCandidateSpacing_(String(value || '').trim());
  if (!text) return '';
  text = stripInlineEditionBrackets_(text);
  text = String(text || '')
    .replace(/\s+(?:資料夾|资料夹|文件夾|文件夹)\s*[\d０-９]{0,4}\s*入\s*$/u, '')
    .trim();
  text = stripLookupVolumeNoise_(text);
  text = stripMangaUpdatesEditionSuffix_(text);
  return normalizeTitleSearchCandidateSpacing_(text);
}

/** rowData の原題·日本語タイトル系をシート向けに整形 */
function finalizeSheetWorkTitleColumns_(rowData) {
  var next = Object.assign({}, rowData || {});
  var jp = String(next['日本語タイトル'] || '').trim();
  if (jp && !isJapaneseTitleLookupStatusValue_(jp)) {
    next['日本語タイトル'] = sanitizeWorkNameJapaneseForSheet_(jp);
  }
  ['作品名（日本語）', '作品名日本語', '商品名（日本語）', '商品名(日本語)'].forEach(function(key) {
    var t = String(next[key] || '').trim();
    if (t && !isJapaneseTitleLookupStatusValue_(t)) {
      next[key] = sanitizeWorkNameJapaneseForSheet_(t);
    }
  });
  [  '原題タイトル', '作品名（原題）'].forEach(function(key) {
    var o = String(next[key] || '').trim();
    if (o) next[key] = sanitizeWorkNameOriginalForSheet_(o);
  });
  /* 雑誌名はマスター入力規則列のためここでは触らない（mag_normalizeMagazineRowForDropdownWrite_ / mag_書込後処理） */
  return next;
}

function buildDictionaryTitleKeys_(value, language, category, originalProductTitle) {
  return uniqNonEmptyTitles_(
    buildTitleSearchCandidates_(language, category, value, originalProductTitle).map(normalizeDictionaryTitleKey_)
  );
}

function buildMangaUpdatesTitleKeys_(value, language, category, originalProductTitle) {
  return uniqNonEmptyTitles_(
    buildTitleSearchCandidates_(language, category, value, originalProductTitle).map(normalizeMangaUpdatesTitleKey_)
  );
}

function cjkCharOverlap_(a, b) {
  var charsA = String(a || '').replace(/[^\u3000-\u9fff\uf900-\ufaff]/g, '');
  var charsB = String(b || '').replace(/[^\u3000-\u9fff\uf900-\ufaff]/g, '');
  if (!charsA || !charsB) return 0;
  var setA = {};
  var setB = {};
  var sizeA = 0;
  var sizeB = 0;
  for (var i = 0; i < charsA.length; i++) { if (!setA[charsA[i]]) { setA[charsA[i]] = true; sizeA++; } }
  for (var j = 0; j < charsB.length; j++) { if (!setB[charsB[j]]) { setB[charsB[j]] = true; sizeB++; } }
  var common = 0;
  for (var ch in setA) { if (setB[ch]) common++; }
  return common / Math.min(sizeA, sizeB);
}

function isStrongMangaUpdatesMatch_(query, searchResult, detail, language, category, originalProductTitle) {
  const queryKeys = buildMangaUpdatesTitleKeys_(query, language, category, originalProductTitle);
  if (!queryKeys.length) return false;

  const associated = Array.isArray(detail && detail.associated) ? detail.associated : [];
  const candidates = uniqNonEmptyTitles_([
    searchResult && searchResult.hit_title,
    searchResult && searchResult.record ? searchResult.record.title : '',
    detail && detail.title,
  ].concat(associated.map(function(item) {
    return item && item.title ? item.title : '';
  })));

  return candidates.some(function(candidate) {
    // CJK文字の重複率でマッチ判定（繁体字↔簡体字↔日本語対応）
    if (cjkCharOverlap_(query, candidate) >= 0.5) return true;
    var candidateKeys = buildMangaUpdatesTitleKeys_(candidate);
    return candidateKeys.some(function(candidateKey) {
      return queryKeys.some(function(queryKey) {
        if (!candidateKey || !queryKey) return false;
        if (candidateKey === queryKey) return true;
        if (candidateKey.length >= 4 && queryKey.length >= 4) {
          return candidateKey.indexOf(queryKey) !== -1 || queryKey.indexOf(candidateKey) !== -1;
        }
        return false;
      });
    });
  });
}

function pickJapaneseMangaUpdatesTitle_(searchResult, detail) {
  const associated = Array.isArray(detail && detail.associated) ? detail.associated : [];
  const candidates = uniqNonEmptyTitles_([
    searchResult && searchResult.hit_title,
    detail && detail.title,
    searchResult && searchResult.record ? searchResult.record.title : '',
  ].concat(associated.map(function(item) {
    return item && item.title ? item.title : '';
  })));

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    if (!hasJapaneseTitleSignal_(candidate)) continue;
    return String(candidate || '').trim();
  }

  return '';
}

function hasJapaneseTitleSignal_(value) {
  return /[ぁ-ゖァ-ヺ々〆ヵヶ]/.test(String(value || ''));
}

function normalizeMangaUpdatesTitleKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000\"'“”‘’´・･:：!！?？,，、.．\-ー‐―–—~〜/／\\|()\[\]{}【】「」『』<>]/g, '');
}

function uniqNonEmptyTitles_(values) {
  const seen = {};
  const list = [];
  (values || []).forEach(function(value) {
    const text = String(value || '').trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    list.push(text);
  });
  return list;
}

function doGet() {
  return jsonResponse_({
    ok: true,
    message: 'books.com.tw web app',
    spreadsheetId: SPREADSHEET_ID,
    masterSpreadsheetId: MASTER_SPREADSHEET_ID,
    comicBookSheetName: COMIC_BOOK_SHEET_NAME,
    otherBookSheetName: OTHER_BOOK_SHEET_NAME,
    goodsSheetName: GOODS_SHEET_NAME,
    magazineSheetName: MAGAZINE_SHEET_NAME,
    magazineMasterSheet: MAGAZINE_MASTER_SHEET,
    magazineMasterCandidateSheet: MAGAZINE_MASTER_CANDIDATE_SHEET,
    editionRuleSheet: MAGAZINE_EDITION_RULE_SHEET,
    typeRuleSheet: MAGAZINE_TYPE_RULE_SHEET,
  });
}

function doGet(e) {
  // ウォームアップ用エンドポイント（GASコールドスタート防止）
  // 拡張機能がポップアップ起動時にここへGETリクエストを送り、
  // 実際のdoPostが呼ばれる前にGASとスプレッドシートを起動しておく。
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (_) {}
  return ContentService.createTextOutput(JSON.stringify({ ok: true, warmed: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 時間トリガー用ウォームアップ関数
 * GASエディタで「トリガーを追加」→ warmup → 時間ベース → 5分ごと に設定してください。
 * これによりGASが常にウォームアップ状態を保ち、コールドスタートを防止します。
 * ※ GASでは末尾に _ がつく関数はプライベート扱いでトリガー登録不可のため warmup_ → warmup に変更
 */
function warmup() {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    Logger.log('warmup error: ' + e.message);
  }
}

function lookupProviderName_(source) {
  const text = String(source || '').trim();
  const providerId = normalizeProviderId_(text);
  if (TITLE_LOOKUP_PROVIDERS[providerId]) return providerId;
  if (text === '辞書') return 'titleAliasDictionary';
  if (text === 'BL補助') return 'chilchil';
  if (text === '候補') return 'candidate';
  return text || '';
}

function japaneseTitleLookupResolvedScore_(source) {
  const id = normalizeProviderId_(source);
  if (id === 'workTitleMaster') return 1100;
  if (id === 'titleAliasDictionary') return 1000;
  return 800;
}

function normalizeLookupWorkTitle_(value) {
  let text = normalizeTitleSearchCandidateSpacing_(String(value || ''));
  if (!text) return '';

  text = stripInlineEditionBrackets_(text);
  text = stripCategorySpecificTitleNoise_(text, 'グッズ');
  text = stripLookupVolumeNoise_(text);
  text = stripMangaUpdatesEditionSuffix_(text);
  text = stripLookupWorkTitleDecoration_(text);
  text = stripMangaUpdatesSubtitle_(text);
  text = normalizeTitleSearchCandidateSpacing_(text);
  return text;
}

function lookupCategoriesForTitleAnalysis_(titleAnalysis, item) {
  const analysis = titleAnalysis || {};
  const itemType = String(analysis.itemType || item && item.itemType || '').trim();
  const rowData = extractRowData_(item || {});
  const rawCategory = analysis.category || rowData['カテゴリ'] || '';
  const categories = [];

  if (itemType === 'manga' || itemType === 'bl_manga') categories.push('まんが');
  if (itemType === 'goods') categories.push('グッズ', 'まんが');
  if (itemType === 'light_novel' || itemType === 'novel_book') categories.push('書籍', 'まんが');
  if (itemType === 'magazine') categories.push('雑誌');
  const normalized = normalizeDictionaryCategory_(rawCategory);
  if (normalized) categories.push(normalized);
  if (!categories.length) categories.push(itemType === 'goods' ? 'グッズ' : '書籍');

  return categories.filter(function(category, index, array) {
    return category && array.indexOf(category) === index;
  });
}

function lookupSearchKeysFromPayload_(payload) {
  const analysis = payload && payload.titleAnalysis || {};
  const rawItem = payload && (payload.rawItem || payload) || {};
  const rowData = extractRowData_(rawItem);
  const keys = [
    analysis.normalizedSearchTitle,
    analysis.extractedWorkTitle,
    analysis.originalTitle,
    analysis.originalProductTitle,
    rowData['原題タイトル'],
    rowData['原題商品タイトル'],
    rowData['原題商品名'],
    rowData['タイトル'],
    rawItem.title,
    payload && payload.title
  ];

  const expandedKeys = [];
  keys.forEach(function(value) {
    const raw = String(value || '').trim();
    const pure = normalizeLookupWorkTitle_(raw);
    if (pure) expandedKeys.push(pure);
    if (raw) expandedKeys.push(raw);
  });

  return expandedKeys.map(function(value) {
    return String(value || '').trim();
  }).filter(function(value, index, array) {
    return value && array.indexOf(value) === index;
  });
}

function normalizeLookupResult_(lookup) {
  const source = lookup && lookup.lookup ? lookup.lookup : lookup;
  if (!source) return null;
  return {
    status: source.status || (source.japaneseTitle || source.title ? 'resolved' : 'not_found'),
    japaneseTitle: source.japaneseTitle || source.title || '',
    provider: source.provider || source.source || '',
    normalizedSearchTitle: source.normalizedSearchTitle || '',
    extractedWorkTitle: source.extractedWorkTitle || '',
    score: source.score || 0,
    trace: source.trace || source.log || '',
    candidates: Array.isArray(source.candidates) ? source.candidates : [],
    errors: Array.isArray(source.errors) ? source.errors : []
  };
}

function lookupJapaneseTitleFromPayload_(payload, skipExternalApi) {
  const analysis = payload && payload.titleAnalysis || {};
  const rawItem = payload && (payload.rawItem || payload) || {};
  const rowData = extractRowData_(rawItem);
  const language = normalizeDictionaryLanguage_(analysis.language || rowData['言語'] || '台湾') || '台湾';
  const categories = lookupCategoriesForTitleAnalysis_(analysis, rawItem);
  const searchKeys = lookupSearchKeysFromPayload_(payload);
  const author = String(analysis.author || rowData['著者'] || rowData['作者'] || '').trim();
  const originalProductTitle = String(analysis.originalProductTitle || rowData['原題商品タイトル'] || rowData['原題商品名'] || rawItem.title || '').trim();
  const trace = [];
  const errors = [];
  let aggregate = {
    checkedSources: [],
    failedSources: [],
    skippedSources: [],
    candidates: [],
    errors: []
  };

  for (let keyIndex = 0; keyIndex < searchKeys.length; keyIndex += 1) {
    const searchKey = searchKeys[keyIndex];
    for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex += 1) {
      const category = categories[categoryIndex];
      try {
        const result = cascadeLookupJapaneseTitleResult_(
          language,
          category,
          searchKey,
          author,
          originalProductTitle,
          analysis.itemType,
          analysis.providerHint,
          skipExternalApi
        );
        aggregate.checkedSources = aggregate.checkedSources.concat(result.checkedSources || []);
        aggregate.failedSources = aggregate.failedSources.concat(result.failedSources || []);
        aggregate.skippedSources = aggregate.skippedSources.concat(result.skippedSources || []);
        aggregate.candidates = aggregate.candidates.concat(result.candidates || []);
        aggregate.errors = aggregate.errors.concat(result.errors || []);
        if (result.japaneseTitle) {
          const provider = lookupProviderName_(result.source);
          const resolvedScore = japaneseTitleLookupResolvedScore_(result.source);
          if (result.trace && result.trace.length) {
            trace.push(searchKey + '=>' + result.trace.join(' > '));
          }
          trace.push((provider || 'provider') + ':hit(key=' + searchKey + ')');
          return {
            success: true,
            lookup: {
              status: 'resolved',
              japaneseTitle: result.japaneseTitle,
              provider: provider,
              normalizedSearchTitle: searchKeys[0] || searchKey,
              extractedWorkTitle: analysis.extractedWorkTitle || searchKey,
              score: resolvedScore,
              trace: trace.join(' | '),
              candidates: [{ title: result.japaneseTitle, provider: provider, score: resolvedScore }],
              errors: errors
            }
          };
        }
        trace.push(result.trace && result.trace.length
          ? searchKey + '=>' + result.trace.join(' > ')
          : 'miss:' + category + ':' + searchKey);
      } catch (error) {
        const message = error && error.message ? error.message : String(error || 'lookup error');
        errors.push(message);
        const failedSource = inferLookupFailureSource_(error);
        if (failedSource) aggregate.failedSources.push(failedSource);
        trace.push('error:' + (failedSource || 'provider') + ':' + searchKey);
      }
    }
  }

  aggregate.checkedSources = uniqNonEmptyTitles_(aggregate.checkedSources);
  aggregate.failedSources = uniqNonEmptyTitles_(aggregate.failedSources);
  aggregate.skippedSources = uniqNonEmptyTitles_(aggregate.skippedSources);
  errors.push.apply(errors, uniqNonEmptyTitles_(aggregate.errors));
  const providerTrace = [];
  if (aggregate.checkedSources.length) providerTrace.push('checked:' + aggregate.checkedSources.join('/'));
  if (aggregate.failedSources.length) providerTrace.push('failed:' + aggregate.failedSources.join('/'));
  if (aggregate.skippedSources.length) providerTrace.push('skipped:' + aggregate.skippedSources.join('/'));
  const candidateTitles = uniqNonEmptyTitles_(aggregate.candidates).slice(0, 5);
  if (candidateTitles.length) providerTrace.push('candidates:' + candidateTitles.join('/'));
  return {
    success: true,
    lookup: {
      status: aggregate.failedSources.length ? 'partial_error' : 'not_found',
      japaneseTitle: '',
      provider: '',
      normalizedSearchTitle: searchKeys[0] || '',
      extractedWorkTitle: analysis.extractedWorkTitle || searchKeys[0] || '',
      score: 0,
      trace: trace.concat(providerTrace).join(' | ') || buildJapaneseTitleLookupStatusLabel_(aggregate),
      candidates: uniqNonEmptyTitles_(aggregate.candidates).map(function(title) {
        return { title: title, provider: 'candidate', score: 100 };
      }),
      errors: errors
    }
  };
}

function lookupJapaneseTitleAction_(payload) {
  return lookupJapaneseTitleFromPayload_(payload || {});
}

// 拡張機能（クライアント）側のカスケードで確定済みとみなせる状態。
// これらが来たら、書き込み時にサーバ側で重い外部API・スクレイピングを
// 再実行しない（＝GASへの書き込みを超高速化する主目的）。
// クライアントの「追加」時点で lookupJapaneseTitle（GASカスケード＋MU直叩き）を
// 既に実行済みのため、書き込み時の再照会は完全な重複処理だった。
const CLIENT_TERMINAL_LOOKUP_STATUSES_ = {
  resolved: true,
  not_found: true,
  series_found_no_japanese: true,
  partial_error: true,
  skipped: true,
};

function shouldRefreshLookup_(lookup, titleAnalysis) {
  const normalized = normalizeLookupResult_(lookup);
  // クライアントが照会結果を一切持っていない場合のみ、サーバ側カスケードを実行。
  if (!normalized) return true;
  const status = String(normalized.status || '').trim();
  // failed / skipped / 未知ステータス（=クライアント側カスケードが完走していない一過性の状態）
  // のときだけサーバで再照会する。
  if (!CLIENT_TERMINAL_LOOKUP_STATUSES_[status]) return true;
  // 確定状態でも、保存済み照会キーが現在のタイトル解析キーと食い違う（陳腐化）場合は再照会。
  const lookupKey = String(normalized.normalizedSearchTitle || '').trim();
  const analysisKey = String(titleAnalysis && titleAnalysis.normalizedSearchTitle || '').trim();
  return Boolean(analysisKey && lookupKey && lookupKey !== analysisKey);
}

/** 拡張の MangaUpdates 直叩きなど、サーバ側カスケードに無い trace 断片をシートの照会ログに残す */
function mergeExtensionMuTraceIntoLookup_(clientLookup, serverLookup) {
  const client = normalizeLookupResult_(clientLookup);
  const server = normalizeLookupResult_(serverLookup);
  if (!server) return client;
  if (!client || !String(client.trace || '').trim()) return server;
  const extraParts = String(client.trace || '')
    .split('|')
    .map(function(p) { return String(p || '').trim(); })
    .filter(function(p) {
      return p && (/mangaupdatesClient:/i.test(p) || /^mangaupdates:/i.test(p));
    });
  if (!extraParts.length) return server;
  const extra = extraParts.join(' | ');
  var sTrace = String(server.trace || '').trim();
  if (!sTrace) return Object.assign({}, server, { trace: extra });
  if (sTrace.indexOf(extra) >= 0) return server;
  return Object.assign({}, server, { trace: sTrace + ' | ' + extra });
}

function japaneseTitleLookupStatusLabel_(lookup) {
  const result = normalizeLookupResult_(lookup);
  if (!result) return JAPANESE_TITLE_NOT_LOOKED_UP_LABEL;
  if (result.status === 'resolved') return 'resolved';
  if (result.status === 'series_found_no_japanese') return 'MU登録あり/日本語訳なし';
  if (result.status === 'not_found') return JAPANESE_TITLE_NOT_FOUND_ALL_SOURCES_LABEL;
  if (result.status === 'partial_error') return '一部照会失敗';
  if (result.status === 'failed') return JAPANESE_TITLE_LOOKUP_FAILED_LABEL + '(' + (result.provider || '不明') + ')';
  if (result.status === 'skipped') return JAPANESE_TITLE_NOT_LOOKED_UP_LABEL;
  return result.status || JAPANESE_TITLE_NOT_LOOKED_UP_LABEL;
}

function applyJapaneseTitleLookupToRowData_(rowData, lookup, titleAnalysis) {
  const result = normalizeLookupResult_(lookup);
  const next = Object.assign({}, rowData || {});
  if (!result) return finalizeSheetWorkTitleColumns_(next);

  if (result.status === 'resolved' && result.japaneseTitle) {
    next['日本語タイトル'] = result.japaneseTitle;
    if (!String(next['作品名（日本語）'] || '').trim()) next['作品名（日本語）'] = result.japaneseTitle;
    if (!String(next['作品名日本語'] || '').trim()) next['作品名日本語'] = result.japaneseTitle;
  } else if (isJapaneseTitleLookupStatusValue_(next['日本語タイトル'])) {
    next['日本語タイトル'] = '';
  }
  next['日本語タイトル照会結果'] = japaneseTitleLookupStatusLabel_(result);
  next['日本語タイトル照会元'] = result.provider || '';
  next['日本語タイトル照会ログ'] = [result.trace || '', (result.errors || []).join(' / ')].filter(Boolean).join(' / ');
  next['検索用正規化タイトル'] = result.normalizedSearchTitle || titleAnalysis && titleAnalysis.normalizedSearchTitle || '';
  return finalizeSheetWorkTitleColumns_(next);
}

function prepareItemWithJapaneseTitleLookup_(item) {
  const titleAnalysis = item && item.titleAnalysis || {};
  const clientLookup = normalizeLookupResult_(item && item.japaneseTitleLookup);
  const clientResolved = !!(clientLookup && clientLookup.status === 'resolved' && String(clientLookup.japaneseTitle || '').trim());
  const clientSeriesOnly = !!(clientLookup && clientLookup.status === 'series_found_no_japanese');

  let lookup;
  if (clientResolved) {
    lookup = clientLookup;
  } else if (shouldRefreshLookup_(item && item.japaneseTitleLookup, titleAnalysis)) {
    lookup = lookupJapaneseTitleFromPayload_({
      source: item && item.source || 'books_tw',
      rawItem: item && (item.rawItem || item),
      titleAnalysis: titleAnalysis
    }, true).lookup;
    lookup = mergeExtensionMuTraceIntoLookup_(item && item.japaneseTitleLookup, lookup);
    if (clientLookup && String(clientLookup.japaneseTitle || '').trim() && !String(lookup.japaneseTitle || '').trim()) {
      lookup = Object.assign({}, lookup, {
        japaneseTitle: clientLookup.japaneseTitle,
        status: 'resolved',
        provider: clientLookup.provider || lookup.provider || 'mangaUpdates(extension)',
      });
    }
    // 「MU上に作品はあるが日本語訳なし」の状態をクライアントから引き継ぐ
    // （サーバカスケード再実行で 'not_found' になっても上書きしない）
    if (clientSeriesOnly && !String(lookup.japaneseTitle || '').trim()) {
      const muCandidates = Array.isArray(clientLookup.candidates) ? clientLookup.candidates : [];
      const existingCandidates = Array.isArray(lookup.candidates) ? lookup.candidates : [];
      lookup = Object.assign({}, lookup, {
        status: 'series_found_no_japanese',
        provider: clientLookup.provider || lookup.provider || 'mangaUpdates(series_no_jp)',
        candidates: uniqNonEmptyTitles_(existingCandidates.concat(muCandidates)),
      });
    }
  } else {
    lookup = clientLookup;
  }
  const rowData = applyJapaneseTitleLookupToRowData_(extractRowData_(item), lookup, titleAnalysis);
  return Object.assign({}, item, {
    rowData: rowData,
    japaneseTitleLookup: lookup
  });
}

function ensureLookupColumns_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const normalized = {};
  headers.forEach(function(header) {
    const key = normalizeHeader_(header);
    if (key) normalized[key] = true;
  });
  const missing = LOOKUP_WRITEBACK_HEADERS.filter(function(header) {
    return !normalized[normalizeHeader_(header)];
  });
  if (!missing.length) return false;
  sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
  return true;
}

function findExistingDuplicateRow_(sheet, duplicateKeyColumns, duplicateKeys) {
  for (let i = 0; i < duplicateKeys.length; i += 1) {
    const duplicateKey = duplicateKeys[i];
    const separatorIndex = String(duplicateKey).indexOf(':');
    if (separatorIndex <= 0) continue;
    const type = duplicateKey.slice(0, separatorIndex);
    const value = duplicateKey.slice(separatorIndex + 1);
    const targetColumn = duplicateKeyColumns.find(function(column) {
      return column && column.type === type;
    });
    if (!targetColumn || !value) continue;
    const pattern = buildDuplicateKeySearchPattern_(type, value);
    if (!pattern) continue;
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow <= 1) continue;
    const finder = sheet
      .getRange(2, targetColumn.index + 1, lastRow - 1, 1)
      .createTextFinder(pattern)
      .useRegularExpression(true)
      .matchCase(false);
    const range = finder.findNext();
    if (range) return range.getRow();
  }
  return 0;
}

/** 1行ぶんをセル単位ではなくまとめて setValues（高速化の主因） */
function writeRowBulk_(context, rowNumber, item) {
  const row = buildRow_(context, item);
  applyTextFormatToColumns_(context.sheet, rowNumber, 1, context.textColumnIndexes);
  // 第3引数は numRows(=1) であって rowNumber ではない。
  // 以前は getRange(rowNumber, 1, rowNumber, lastColumn) としていたため
  // 行2以降で「データ次元不一致」となり setValues が失敗、結果的に
  // 日本語タイトル照会ログ等の列が更新されない不具合の原因になっていた。
  const rng = context.sheet.getRange(rowNumber, 1, 1, context.lastColumn);
  rng.setValues([row]);
  rng.setWrap(false);
  applyRichTextLinksToColumns_(context.sheet, rowNumber, [row], context.urlLinkColumnIndexes);
}

/**
 * 既存行の重複キー → 行番号（1スキャンで構築。TextFinder 繰り返しより高速）
 */
function buildDuplicateKeyToRowMapInMemory_(allValues, duplicateKeyColumns) {
  const map = Object.create(null);
  if (!duplicateKeyColumns || !duplicateKeyColumns.length) return map;

  const numRows = allValues.length;
  if (numRows <= 0) return map;

  duplicateKeyColumns.forEach(function(column) {
    const colIdx = column.index;
    for (let r = 0; r < numRows; r += 1) {
      const cellRaw = allValues[r][colIdx];
      const cellStr = String(cellRaw == null ? '' : cellRaw);
      const normalized = normalizeDuplicateKeyValueByType_(column.type, cellStr);
      if (!normalized) continue;
      map[column.type + ':' + normalized] = r + 2;
    }
  });
  return map;
}

function buildDuplicateKeyToRowMap_(sheet, duplicateKeyColumns, sampleItem) {
  const map = Object.create(null);
  if (!duplicateKeyColumns || !duplicateKeyColumns.length) return map;

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headerValues = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const normalizedHeaderMap = {};
  headerValues.forEach(function(header, idx) {
    const normalized = normalizeHeader_(header);
    if (normalized) normalizedHeaderMap[normalized] = idx;
  });
  const rowData = sampleItem ? extractRowData_(sampleItem) : {};
  const lastRow = findActualLastDataRow_(sheet, normalizedHeaderMap, sampleItem, rowData);
  if (lastRow <= 1) return map;

  const colIndexes = duplicateKeyColumns.map(function(c) { return c.index; });
  const minCol = Math.min.apply(null, colIndexes);
  const maxCol = Math.max.apply(null, colIndexes);
  const numRows = lastRow - 1;
  const allValues = sheet.getRange(2, minCol + 1, lastRow, maxCol + 1).getDisplayValues();

  duplicateKeyColumns.forEach(function(column) {
    const localCol = column.index - minCol;
    for (let r = 0; r < numRows; r += 1) {
      const normalized = normalizeDuplicateKeyValueByType_(column.type, allValues[r][localCol]);
      if (!normalized) continue;
      map[column.type + ':' + normalized] = r + 2;
    }
  });
  return map;
}

function findExistingRowFromKeyMaps_(duplicateKeys, primaryMap, fallbackMap) {
  if (!duplicateKeys || !duplicateKeys.length) return 0;
  for (let i = 0; i < duplicateKeys.length; i += 1) {
    const direct = primaryMap[duplicateKeys[i]];
    if (direct) return direct;
    const pend = fallbackMap && fallbackMap[duplicateKeys[i]];
    if (pend) return pend;
  }
  return 0;
}

function upsertItemsWithLookup_(ss, items, timing) {
  timing = timing || {};
  
  var tEnrichStart = Date.now();
  var preparedEntries = [];
  var ix;
  for (ix = 0; ix < items.length; ix += 1) {
    preparedEntries.push({
      index: ix,
      prepared: mag_normalizeMagazineRowForDropdownWrite_(prepareItemWithJapaneseTitleLookup_(items[ix])),
    });
  }
  timing.enrichMs = Date.now() - tEnrichStart;

  var bySheet = {};
  for (ix = 0; ix < preparedEntries.length; ix += 1) {
    var e = preparedEntries[ix];
    var sname = resolveSheetName_(e.prepared);
    e.sheetName = sname;
    if (!bySheet[sname]) bySheet[sname] = [];
    bySheet[sname].push(e);
  }

  var results = new Array(items.length);
  Object.keys(bySheet).forEach(function(sheetName) {
    var group = bySheet[sheetName];
    var samplePrepared = group[0].prepared;

    var tBStart = Date.now();
    var context = buildSheetContext_(ss, sheetName, samplePrepared);
    // 照会ログ等の列が不足しているときだけヘッダー追記＆コンテキスト再構築する。
    // 既に列が揃っている通常時は再読込しない（無駄なシート読込を削減）。
    var columnsAdded = ensureLookupColumns_(context.sheet);
    var refreshedContext = columnsAdded
      ? buildSheetContext_(ss, sheetName, samplePrepared)
      : context;
    timing.buildContextMs = (timing.buildContextMs || 0) + (Date.now() - tBStart);

    var tKStart = Date.now();
    var duplicateKeyColumns = findDuplicateKeyColumns_(refreshedContext.normalizedHeaderMap, samplePrepared);
    var keyRowOnSheet = buildDuplicateKeyToRowMapInMemory_(refreshedContext.allValues, duplicateKeyColumns);
    var pendingKeyRow = Object.create(null);
    timing.buildKeysMs = (timing.buildKeysMs || 0) + (Date.now() - tKStart);

    var gj;
    for (gj = 0; gj < group.length; gj += 1) {
      var entry = group[gj];
      var preparedItem = entry.prepared;
      var rowData = extractRowData_(preparedItem);
      var duplicateKeys = buildDuplicateKeysForRowData_(rowData, duplicateKeyColumns);
      // メモリ上のキーマップのみ（事前にシート全域を読み済み）。新規時に TextFinder で列再走査しない
      var existingRow = findExistingRowFromKeyMaps_(duplicateKeys, keyRowOnSheet, pendingKeyRow);

      var targetRow = existingRow;
      var mode = 'updated';

      var tWStart = Date.now();
      if (!targetRow) {
        // 事前に読み込んだ allValues から追記行を決定（シート再読込を避けて高速化）。
        targetRow = resolveAppendRowsInMemory_(refreshedContext.allValues, refreshedContext.keyColumnIndex, 1)[0];
        mode = 'created';
        var newRow = buildRow_(refreshedContext, preparedItem);
        writeRows_(refreshedContext.sheet, refreshedContext.lastColumn, [targetRow], [newRow], refreshedContext.textColumnIndexes, refreshedContext.urlLinkColumnIndexes);
      } else {
        writeRowBulk_(refreshedContext, targetRow, preparedItem);
      }
      timing.writeRowsMs = (timing.writeRowsMs || 0) + (Date.now() - tWStart);

      var pk2;
      for (pk2 = 0; pk2 < duplicateKeys.length; pk2 += 1) {
        pendingKeyRow[duplicateKeys[pk2]] = targetRow;
        keyRowOnSheet[duplicateKeys[pk2]] = targetRow;
      }

      if (sheetName === MAGAZINE_SHEET_NAME) {
        try {
          mag_書込後処理_(refreshedContext.sheet, targetRow, preparedItem);
        } catch (err) {
          Logger.log('台湾雑誌 書込後処理エラー row ' + targetRow + ': ' + err.message);
        }
      }

      results[entry.index] = {
        ok: true,
        success: true,
        index: entry.index,
        sheetName: sheetName,
        productCode: preparedItem && preparedItem.productCode ? preparedItem.productCode : '',
        row: targetRow,
        appendedRow: targetRow,
        mode: mode,
        lookup: preparedItem.japaneseTitleLookup || null,
        sheetRows: refreshedContext.allValues.length + 1,
        sheetCols: refreshedContext.lastColumn
      };
    }
  });
  timing.writeEnd = Date.now();
  return results;
}

function upsertProductWithLookupAction_(payload) {
  const items = Array.isArray(payload && payload.items)
    ? payload.items
    : [payload && Object.assign({}, payload.rawItem || {}, {
      source: payload.source || 'books_tw',
      titleAnalysis: payload.titleAnalysis || null,
      japaneseTitleLookup: payload.japaneseTitleLookup || null
    })].filter(Boolean);
  if (!items.length) throw new Error('items is empty');

  // === 処理時間計測（書き込みが遅いときの切り分け用）===
  const tStart = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const tLock = Date.now();
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const serverTiming = {
      lockWaitMs: tLock - tStart,
    };
    if (payload.compactSheets !== false) {
      const tCompactStart = Date.now();
      serverTiming.compactSheets = compactSheetsForItems_(ss, items);
      serverTiming.compactMs = Date.now() - tCompactStart;
    }
    const results = upsertItemsWithLookup_(ss, items, serverTiming);
    const tEnd = Date.now();
    const first = results[0] || {};
    return {
      success: true,
      ok: true,
      mode: first.mode || '',
      source: payload && payload.source || 'books_tw',
      sheet: first.sheetName || '',
      sheetName: first.sheetName || '',
      spreadsheetId: SPREADSHEET_ID,
      row: first.row || first.appendedRow || '',
      lookup: first.lookup || null,
      sheetRows: first.sheetRows || 0,
      sheetCols: first.sheetCols || 0,
      results: results,
      // 各工程の実測ミリ秒を返す
      timing: {
        serverMs: tEnd - tStart,
        lockWaitMs: serverTiming.lockWaitMs,
        compactMs: serverTiming.compactMs || 0,
        enrichMs: serverTiming.enrichMs,
        buildContextMs: serverTiming.buildContextMs,
        buildKeysMs: serverTiming.buildKeysMs,
        writeRowsMs: serverTiming.writeRowsMs,
        recalcMs: tEnd - tLock - (serverTiming.compactMs || 0) - (serverTiming.enrichMs || 0) - (serverTiming.buildContextMs || 0) - (serverTiming.buildKeysMs || 0) - (serverTiming.writeRowsMs || 0),
      },
      gasFile: "博客來ウェブアプリ.js",
      gasVersion: "v73",
      warnings: []
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doPost(e) {
  const raw = e && e.postData ? e.postData.contents : '';

  try {
    if (!raw) throw new Error('postData.contents is empty');

    const payload = JSON.parse(raw);
    if (payload && payload.action === 'lookupJapaneseTitle') {
      return jsonResponse_(lookupJapaneseTitleAction_(payload));
    }
    if (payload && payload.action === 'upsertProductWithLookup') {
      return jsonResponse_(upsertProductWithLookupAction_(payload));
    }
    if (payload && payload.action === 'lookupMangaUpdatesJapaneseTitle') {
      return handleMangaUpdatesLookupPost_(payload);
    }
    if (payload && payload.action === 'compactSheets') {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const items = Array.isArray(payload.items) ? payload.items : [];
        const sheetNames = Array.isArray(payload.sheetNames) ? payload.sheetNames : [];
        if (items.length) {
          return jsonResponse_({ ok: true, compacted: compactSheetsForItems_(ss, items) });
        }
        const trimmed = {};
        sheetNames.forEach(function(sheetName) {
          const sheet = ss.getSheetByName(String(sheetName || '').trim());
          if (!sheet) return;
          trimmed[sheetName] = compactSheetForWrite_(sheet, null);
        });
        return jsonResponse_({ ok: true, compacted: trimmed });
      } finally {
        try { lock.releaseLock(); } catch (_) {}
      }
    }
    const items = Array.isArray(payload.items) ? payload.items : [];
    if (!items.length) throw new Error('items is empty');

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const results = appendItems_(ss, items);
      const appended = results.filter(function(r) { return r && r.ok && !r.skipped; }).length;
      const skipped = results.filter(function(r) { return r && r.skipped; }).length;

      return jsonResponse_({
        ok: true,
        appended: appended,
        skipped: skipped,
        results: results,
        source: payload.source || 'unknown',
        spreadsheetId: SPREADSHEET_ID,
        masterSpreadsheetId: typeof MASTER_SPREADSHEET_ID !== 'undefined' ? MASTER_SPREADSHEET_ID : '',
        bookSheetName: typeof BOOK_SHEET_NAME !== 'undefined' ? BOOK_SHEET_NAME : '',
        comicBookSheetName: typeof COMIC_BOOK_SHEET_NAME !== 'undefined' ? COMIC_BOOK_SHEET_NAME : '',
        otherBookSheetName: typeof OTHER_BOOK_SHEET_NAME !== 'undefined' ? OTHER_BOOK_SHEET_NAME : '',
        goodsSheetName: GOODS_SHEET_NAME,
        magazineSheetName: MAGAZINE_SHEET_NAME,
      });
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error.message,
      stack: error.stack || '',
    });
  }
}

function appendItems_(ss, items) {
  const groupedBySheet = {};

  items.forEach((item, index) => {
    const sheetName = resolveSheetName_(item);
    if (!groupedBySheet[sheetName]) groupedBySheet[sheetName] = [];
    groupedBySheet[sheetName].push({ item, index });
  });

  const results = new Array(items.length);

  Object.keys(groupedBySheet).forEach(sheetName => {
    const entries = groupedBySheet[sheetName];
    const context = buildSheetContext_(ss, sheetName, entries[0].item);
    const duplicateKeyColumns = findDuplicateKeyColumns_(context.normalizedHeaderMap, entries[0].item);
    const pendingDuplicateKeys = new Set();
    const appendEntries = [];
    const rows = [];

    entries.forEach(entry => {
      const convertedItem = mag_normalizeMagazineRowForDropdownWrite_(convertJapaneseTitle_(entry.item));
      const duplicateKeys = buildDuplicateKeysForRowData_(extractRowData_(convertedItem), duplicateKeyColumns);
      const duplicateHit = duplicateKeys.find(key => {
        if (pendingDuplicateKeys.has(key)) return true;
        return hasExistingDuplicateKeyInMemory_(context.allValues, duplicateKeyColumns, key);
      });
      if (duplicateHit) {
        results[entry.index] = {
          ok: true,
          skipped: true,
          reason: 'duplicate',
          index: entry.index,
          sheetName,
          duplicateKey: duplicateHit,
          productCode: entry.item && entry.item.productCode ? entry.item.productCode : '',
        };
        return;
      }

      duplicateKeys.forEach(key => pendingDuplicateKeys.add(key));
      appendEntries.push(Object.assign({}, entry, { item: convertedItem }));
      rows.push(buildRow_(context, convertedItem));
    });

    if (!appendEntries.length) return;

    const targetRows = resolveAppendRowsInMemory_(context.allValues, context.keyColumnIndex, rows.length);
    writeRows_(context.sheet, context.lastColumn, targetRows, rows, context.textColumnIndexes, context.urlLinkColumnIndexes);

    appendEntries.forEach((entry, entryIndex) => {
      const appendedRow = targetRows[entryIndex];

      if (sheetName === MAGAZINE_SHEET_NAME) {
        try {
          mag_書込後処理_(context.sheet, appendedRow, entry.item);
        } catch (err) {
          Logger.log(`台湾雑誌 書込後処理エラー row ${appendedRow}: ${err.message}`);
        }
      }

      results[entry.index] = {
        ok: true,
        index: entry.index,
        sheetName,
        productCode: entry.item && entry.item.productCode ? entry.item.productCode : '',
        appendedRow,
      };
    });
  });

  return results;
}

function resolveSheetName_(item) {
  if (item && item.sheetName) return String(item.sheetName);
  if (item && item.itemType === 'goods') return GOODS_SHEET_NAME;
  if (item && item.itemType === 'magazine') return MAGAZINE_SHEET_NAME;

  const rowData = extractRowData_(item);
  const category = String(rowData['カテゴリ'] || '').trim();
  return category === 'まんが' ? COMIC_BOOK_SHEET_NAME : OTHER_BOOK_SHEET_NAME;
}

var DATA_ROW_SCAN_CHUNK = 500;
var SHEET_TRIM_PADDING_ROWS = 25;
var SHEET_TRIM_MIN_SURPLUS_ROWS = 100;

function collectDataScanColumnIndexes_(normalizedHeaderMap, sampleItem, rowData) {
  const indexes = [];
  const keyIdx = findAppendKeyColumnIndex_(normalizedHeaderMap, sampleItem, rowData || {});
  if (typeof keyIdx === 'number' && keyIdx >= 0) indexes.push(keyIdx);

  [
    '原題タイトル', '原題商品タイトル', 'タイトル', 'サイト商品コード', '博客來商品コード',
    '商品コード（SKU）', 'リンク', '重複チェックキー', 'カテゴリ', 'ISBN',
  ].forEach(function(header) {
    const idx = normalizedHeaderMap[normalizeHeader_(header)];
    if (typeof idx === 'number' && idx >= 0 && indexes.indexOf(idx) < 0) indexes.push(idx);
  });

  if (!indexes.length) {
    const fallbackCols = Math.min(3, Math.max(sheet_getLastColumnSafe_(normalizedHeaderMap), 1));
    for (let c = 0; c < fallbackCols; c += 1) indexes.push(c);
  }
  return indexes;
}

function sheet_getLastColumnSafe_(normalizedHeaderMap) {
  const values = Object.keys(normalizedHeaderMap).map(function(key) {
    return normalizedHeaderMap[key];
  });
  if (!values.length) return 1;
  return Math.max.apply(null, values) + 1;
}

function rowHasMeaningfulSheetData_(rowValues, colIndexes, minCol1) {
  for (let i = 0; i < colIndexes.length; i += 1) {
    const localCol = colIndexes[i] - (minCol1 - 1);
    if (localCol < 0 || localCol >= rowValues.length) continue;
    const text = String(rowValues[localCol] == null ? '' : rowValues[localCol]).trim();
    if (text.length >= 2) return true;
  }
  return false;
}

/**
 * getLastRow() が空行膨張で5万超えでも、キー列を下からチャンク読みして真の最終行だけ返す。
 */
function findActualLastDataRow_(sheet, normalizedHeaderMap, sampleItem, rowData) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow <= 1) return 1;

  const colIndexes = collectDataScanColumnIndexes_(normalizedHeaderMap, sampleItem, rowData);
  const minCol1 = Math.min.apply(null, colIndexes) + 1;
  const maxCol1 = Math.max.apply(null, colIndexes) + 1;
  const width = maxCol1 - minCol1 + 1;

  function rowHasData(row) {
    if (row < 2) return false;
    const vals = sheet.getRange(row, minCol1, row, width).getValues()[0];
    return rowHasMeaningfulSheetData_([vals], colIndexes, minCol1);
  }

  // 下側が空行膨張（データは上の方だけ）のとき、下からチャンク走査すると遅いので二分探索する。
  if (lastRow > 1000 && !rowHasData(lastRow)) {
    let lo = 2;
    let hi = lastRow;
    let ans = 1;
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (rowHasData(mid)) {
        ans = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return ans;
  }

  let end = lastRow;
  while (end >= 2) {
    const start = Math.max(2, end - DATA_ROW_SCAN_CHUNK + 1);
    const values = sheet.getRange(start, minCol1, end, width).getValues();
    for (let r = values.length - 1; r >= 0; r -= 1) {
      if (rowHasMeaningfulSheetData_(values[r], colIndexes, minCol1)) {
        return start + r;
      }
    }
    end = start - 1;
  }
  return 1;
}

function maybeTrimSheetEmptyTail_(sheet, actualLastRow) {
  const keepRows = Math.max(2, actualLastRow + SHEET_TRIM_PADDING_ROWS);
  const maxRows = sheet.getMaxRows();
  if (maxRows <= keepRows + SHEET_TRIM_MIN_SURPLUS_ROWS) {
    return { trimmed: 0, keepRows: keepRows };
  }
  const deleteCount = maxRows - keepRows;
  sheet.deleteRows(keepRows + 1, deleteCount);
  return { trimmed: deleteCount, keepRows: keepRows };
}

function compactSheetForWrite_(sheet, sampleItem) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headerValues = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const normalizedHeaderMap = {};
  headerValues.forEach(function(header, idx) {
    const normalized = normalizeHeader_(header);
    if (normalized) normalizedHeaderMap[normalized] = idx;
  });
  const rowData = sampleItem ? extractRowData_(sampleItem) : {};
  const actualLastRow = findActualLastDataRow_(sheet, normalizedHeaderMap, sampleItem, rowData);
  return maybeTrimSheetEmptyTail_(sheet, actualLastRow);
}

function compactSheetsForItems_(ss, items) {
  const names = Object.create(null);
  items.forEach(function(item) {
    names[resolveSheetName_(item)] = true;
  });
  const trimmed = {};
  Object.keys(names).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    trimmed[sheetName] = compactSheetForWrite_(sheet, items.find(function(item) {
      return resolveSheetName_(item) === sheetName;
    }) || null);
  });
  return trimmed;
}

function buildSheetContext_(ss, sheetName, sampleItem) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headerValues = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  if (!headerValues.some(Boolean)) throw new Error(`Header row is empty: ${sheetName}`);

  const normalizedHeaderMap = {};
  headerValues.forEach((header, idx) => {
    const normalized = normalizeHeader_(header);
    if (normalized) normalizedHeaderMap[normalized] = idx;
  });

  const rowData = extractRowData_(sampleItem);
  const actualLastRow = findActualLastDataRow_(sheet, normalizedHeaderMap, sampleItem, rowData);
  if (sheet.getMaxRows() > actualLastRow + SHEET_TRIM_MIN_SURPLUS_ROWS + SHEET_TRIM_PADDING_ROWS) {
    maybeTrimSheetEmptyTail_(sheet, actualLastRow);
  }

  const allValues = actualLastRow > 1
    ? sheet.getRange(2, 1, actualLastRow - 1, lastColumn).getValues()
    : [];

  return {
    sheet,
    sheetName,
    lastColumn,
    normalizedHeaderMap,
    allValues,
    actualLastRow: actualLastRow,
    keyColumnIndex: findAppendKeyColumnIndex_(normalizedHeaderMap, sampleItem, rowData),
    textColumnIndexes: findTextPreservingColumnIndexes_(normalizedHeaderMap),
    urlLinkColumnIndexes: findUrlLinkColumnIndexes_(normalizedHeaderMap),
  };
}

function sanitizeIsbn_(value) {
  const text = String(value || '').replace(/[^\d]/g, '');
  return /^\d{13}$/.test(text) ? text : '';
}

function buildRow_(context, item) {
  const rowData = extractRowData_(item);
  const row = new Array(context.lastColumn).fill('');

  Object.keys(rowData).forEach(key => {
    const idx = context.normalizedHeaderMap[normalizeHeader_(key)];
    if (typeof idx !== 'number') return;
    const normalizedKey = normalizeHeader_(key);
    const rawValue = rowData[key];
    // ISBN列は13桁のみ許可
    const cellValue = (normalizedKey === 'ISBN' || normalizedKey === 'isbn')
      ? sanitizeIsbn_(rawValue)
      : toCellValue_(rawValue);
    row[idx] = cellValue;
  });

  return row;
}


function findDuplicateKeyGroups_(item) {
  if (item && item.itemType === 'goods') {
    return [
      { type: 'productCode', headers: ['博客來商品コード', 'サイト商品コード', '商品コード', 'productCode'] },
      { type: 'duplicateKey', headers: ['重複チェックキー'] },
      { type: 'link', headers: ['博客來URL', 'リンク', 'URL', 'pageUrl'] },
    ];
  }

  if (item && item.itemType === 'magazine') {
    return [
      { type: 'productCode', headers: ['博客來商品コード', 'サイト商品コード', '商品コード', 'productCode'] },
      { type: 'link', headers: ['博客來URL', 'リンク', 'URL', 'pageUrl'] },
    ];
  }

  return [
    { type: 'productCode', headers: ['サイト商品コード', '博客來商品コード', '商品コード', 'productCode'] },
    { type: 'link', headers: ['リンク', '博客來URL', 'URL', 'pageUrl'] },
    { type: 'isbn', headers: ['ISBN', 'isbn'] },
    { type: 'sku', headers: ['商品コード（SKU）'] },
  ];
}

function findDuplicateKeyColumns_(normalizedHeaderMap, item) {
  return findDuplicateKeyGroups_(item).map(group => {
    for (let i = 0; i < group.headers.length; i += 1) {
      const idx = normalizedHeaderMap[normalizeHeader_(group.headers[i])];
      if (typeof idx === 'number') {
        return {
          type: group.type,
          index: idx,
          headers: group.headers,
        };
      }
    }
    return null;
  }).filter(Boolean);
}

function normalizeDuplicateKeyValue_(value) {
  return normalizeCellText_(value).toLowerCase();
}

function normalizeDuplicateKeyValueByType_(type, value) {
  const normalized = normalizeDuplicateKeyValue_(value);
  if (!normalized) return '';

  if (type === 'productCode') {
    const compact = normalized.replace(/\s+/g, '');
    if (/^\d+$/.test(compact)) {
      return compact.replace(/^0+(?=\d)/, '');
    }
  }

  return normalized;
}

function buildDuplicateKeysForRowData_(rowData, duplicateKeyColumns) {
  const keys = [];
  duplicateKeyColumns.forEach(column => {
    let rawValue = '';
    for (let i = 0; i < column.headers.length; i += 1) {
      const candidate = rowData && rowData[column.headers[i]];
      if (candidate == null) continue;
      const normalized = normalizeDuplicateKeyValueByType_(column.type, candidate);
      if (!normalized) continue;
      rawValue = normalized;
      break;
    }
    if (!rawValue) return;
    keys.push(`${column.type}:${rawValue}`);
  });
  return [...new Set(keys)];
}

function collectExistingDuplicateKeys_(sheet, duplicateKeyColumns) {
  const seen = new Set();
  if (!duplicateKeyColumns.length) return seen;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow <= 1) return seen;

  // CacheServiceで重複キーをキャッシュ（5分間）
  const cacheKey = `dupkeys_${sheet.getSheetId()}_${lastRow}`;
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      JSON.parse(cached).forEach(k => seen.add(k));
      return seen;
    }
  } catch (_) {}

  // 全必要列のインデックスを集約し、含む最小～最大範囲を一括読み取り
  const colIndexes = duplicateKeyColumns.map(c => c.index);
  const minCol = Math.min.apply(null, colIndexes);
  const maxCol = Math.max.apply(null, colIndexes);
  const numRows = lastRow - 1;
  const numCols = maxCol - minCol + 1;
  const allValues = sheet.getRange(2, minCol + 1, numRows, numCols).getDisplayValues();

  duplicateKeyColumns.forEach(column => {
    const localCol = column.index - minCol;
    for (let r = 0; r < numRows; r += 1) {
      const normalized = normalizeDuplicateKeyValueByType_(column.type, allValues[r][localCol]);
      if (!normalized) continue;
      seen.add(`${column.type}:${normalized}`);
    }
  });

  try {
    cache.put(cacheKey, JSON.stringify([...seen]), 300);
  } catch (_) {}

  return seen;
}

function hasExistingDuplicateKeyInMemory_(allValues, duplicateKeyColumns, duplicateKey) {
  if (!duplicateKey) return false;

  const separatorIndex = String(duplicateKey).indexOf(':');
  if (separatorIndex <= 0) return false;

  const type = duplicateKey.slice(0, separatorIndex);
  const value = String(duplicateKey.slice(separatorIndex + 1)).trim().toLowerCase();
  if (!value) return false;

  const targetColumn = duplicateKeyColumns.find(function(column) {
    return column && column.type === type;
  });
  if (!targetColumn) return false;

  const colIdx = targetColumn.index;

  for (let r = 0; r < allValues.length; r += 1) {
    const cellRaw = allValues[r][colIdx];
    const cellStr = String(cellRaw == null ? '' : cellRaw).trim().toLowerCase();
    if (!cellStr) continue;

    if (type === 'productCode' && /^\d+$/.test(value)) {
      const valNum = value.replace(/^0+/, '');
      const cellNum = cellStr.replace(/^0+/, '');
      if (valNum === cellNum) return true;
    } else {
      if (cellStr === value) return true;
    }
  }

  return false;
}

function hasExistingDuplicateKey_(sheet, duplicateKeyColumns, duplicateKey) {
  if (!duplicateKey) return false;

  const separatorIndex = String(duplicateKey).indexOf(':');
  if (separatorIndex <= 0) return false;

  const type = duplicateKey.slice(0, separatorIndex);
  const value = duplicateKey.slice(separatorIndex + 1);
  if (!value) return false;

  const targetColumn = duplicateKeyColumns.find(function(column) {
    return column && column.type === type;
  });
  if (!targetColumn) return false;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow <= 1) return false;

  const pattern = buildDuplicateKeySearchPattern_(type, value);
  if (!pattern) return false;

  const finder = sheet
    .getRange(2, targetColumn.index + 1, lastRow - 1, 1)
    .createTextFinder(pattern)
    .useRegularExpression(true)
    .matchCase(false);

  return !!finder.findNext();
}

function buildDuplicateKeySearchPattern_(type, value) {
  const escaped = escapeRegexText_(String(value || ''));
  if (!escaped) return '';
  if (type === 'productCode' && /^\d+$/.test(String(value || ''))) {
    return '^0*' + escaped + '$';
  }
  return '^' + escaped + '$';
}

function escapeRegexText_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findTextPreservingColumnIndexes_(normalizedHeaderMap) {
  const headers = [
    'サイト商品コード',
    '博客來商品コード',
    '商品コード',
    'productCode',
    '商品コード（SKU）',
    'ISBN',
    'isbn',
    'アラジン商品コード',
  ];
  const indexes = new Set();

  headers.forEach(header => {
    const idx = normalizedHeaderMap[normalizeHeader_(header)];
    if (typeof idx === 'number') indexes.add(idx);
  });

  return Array.from(indexes).sort((a, b) => a - b);
}

function findUrlLinkColumnIndexes_(normalizedHeaderMap) {
  const headers = [
    'リンク',
    'URL',
    '博客來URL',
    'メイン画像',
    '追加画像',
    'メイン画像URL',
    '追加画像URL',
    '画像URL',
  ];
  const indexes = new Set();

  headers.forEach(header => {
    const idx = normalizedHeaderMap[normalizeHeader_(header)];
    if (typeof idx === 'number') indexes.add(idx);
  });

  return Array.from(indexes).sort((a, b) => a - b);
}

function extractRowData_(item) {
  return item && item.rowData && typeof item.rowData === 'object' && !Array.isArray(item.rowData)
    ? item.rowData
    : {};
}

function findAppendKeyColumnIndex_(normalizedHeaderMap, item, rowData) {
  let candidates;

  if (item && item.itemType === 'goods') {
    candidates = ['博客來商品コード', 'サイト商品コード', '重複チェックキー', '博客來URL', 'リンク'];
  } else if (item && item.itemType === 'magazine') {
    candidates = ['博客來商品コード', '原題タイトル', '原題商品名', '博客來URL', 'リンク', '雑誌名'];
  } else {
    candidates = ['原題タイトル', '原題商品タイトル', 'タイトル', 'リンク', 'サイト商品コード', '商品コード（SKU）'];
  }

  Object.keys(rowData || {}).forEach(key => {
    if (!candidates.includes(key)) candidates.push(key);
  });

  for (let i = 0; i < candidates.length; i += 1) {
    const idx = normalizedHeaderMap[normalizeHeader_(candidates[i])];
    if (typeof idx === 'number') return idx;
  }

  return -1;
}

function resolveAppendRowsInMemory_(allValues, keyColumnIndex, count) {
  if (count <= 0) return [];

  const dataLastRow = keyColumnIndex >= 0
    ? findLastFilledRowInColumnInMemory_(allValues, keyColumnIndex)
    : allValues.length + 1;

  return buildSequentialRows_(Math.max(dataLastRow, 1) + 1, count);
}

function findLastFilledRowInColumnInMemory_(allValues, colIdx) {
  for (let r = allValues.length - 1; r >= 0; r -= 1) {
    if (String(allValues[r][colIdx] || '').trim()) {
      return r + 2; // 2-indexed
    }
  }
  return 1;
}

function resolveAppendRows_(sheet, keyColumnIndex, count) {
  if (count <= 0) return [];

  const dataLastRow = keyColumnIndex >= 0
    ? findLastFilledRowInColumn_(sheet, keyColumnIndex + 1)
    : Math.max(sheet.getLastRow(), 1);

  return buildSequentialRows_(Math.max(dataLastRow, 1) + 1, count);
}

function findLastFilledRowInColumn_(sheet, columnNumber) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow <= 1) return 1;

  const chunkSize = 500;
  for (let endRow = lastRow; endRow >= 2; endRow -= chunkSize) {
    const startRow = Math.max(2, endRow - chunkSize + 1);
    const rowCount = endRow - startRow + 1;
    const values = sheet.getRange(startRow, columnNumber, rowCount, 1).getDisplayValues();

    for (let i = values.length - 1; i >= 0; i -= 1) {
      if (String(values[i][0] || '').trim()) return startRow + i;
    }
  }

  return 1;
}

function buildSequentialRows_(startRow, count) {
  return Array.from({ length: count }, (_, index) => startRow + index);
}

function applyTextFormatToColumns_(sheet, startRow, rowCount, columnIndexes) {
  if (!rowCount || !columnIndexes || !columnIndexes.length) return;

  // getRangeList でまとめて1回のAPI呼び出しに集約（列ごと個別呼び出しより高速）
  const a1Notations = columnIndexes.map(index =>
    sheet.getRange(startRow, index + 1, rowCount, 1).getA1Notation()
  );
  sheet.getRangeList(a1Notations).setNumberFormat('@');
}

function buildUrlRichTextValue_(value) {
  const text = String(value || '');
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const urlPattern = /https?:\/\/[^\s]+/g;
  let match;
  let hasUrl = false;

  while ((match = urlPattern.exec(text)) !== null) {
    let url = match[0];
    let end = match.index + url.length;
    while (url && /[)\]}>、。，,;；]+$/.test(url)) {
      url = url.slice(0, -1);
      end -= 1;
    }
    if (!url) continue;
    builder.setLinkUrl(match.index, end, url);
    hasUrl = true;
  }

  return hasUrl ? builder.build() : null;
}

function applyRichTextLinksToColumns_(sheet, startRow, rows, columnIndexes) {
  if (!rows || !rows.length || !columnIndexes || !columnIndexes.length) return;

  columnIndexes.forEach(index => {
    const richValues = [];
    let hasAnyUrl = false;
    for (let r = 0; r < rows.length; r += 1) {
      const text = String((rows[r] || [])[index] || '');
      const richValue = buildUrlRichTextValue_(text);
      if (richValue) hasAnyUrl = true;
      richValues.push([richValue || SpreadsheetApp.newRichTextValue().setText(text).build()]);
    }
    if (!hasAnyUrl) return;
    sheet.getRange(startRow, index + 1, rows.length, 1).setRichTextValues(richValues);
  });
}

function writeRows_(sheet, lastColumn, targetRows, rows, textColumnIndexes, urlLinkColumnIndexes) {
  if (!targetRows.length || !rows.length) return;

  const maxTargetRow = Math.max.apply(null, targetRows);
  if (maxTargetRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), maxTargetRow - sheet.getMaxRows());
  }

  let startIndex = 0;
  while (startIndex < targetRows.length) {
    let endIndex = startIndex;
    while (endIndex + 1 < targetRows.length && targetRows[endIndex + 1] === targetRows[endIndex] + 1) {
      endIndex += 1;
    }

    const startRow = targetRows[startIndex];
    const segmentRows = rows.slice(startIndex, endIndex + 1);
    const range = sheet.getRange(startRow, 1, segmentRows.length, lastColumn);
    // テキスト書式を先に適用してからデータ書き込み（順序を保持しつつバッファ活用）
    applyTextFormatToColumns_(sheet, startRow, segmentRows.length, textColumnIndexes);
    range.setValues(segmentRows);
    range.setWrap(false);
    applyRichTextLinksToColumns_(sheet, startRow, segmentRows, urlLinkColumnIndexes);

    startIndex = endIndex + 1;
  }

  // 全セグメント書き込み完了後に一括フラッシュ（API遅延を最小化）
  // ※同期的な計算待ちを防ぐため、コメントアウトして非同期に完了させます
  // SpreadsheetApp.flush();
}

function normalizeHeader_(value) {
  return String(value || '').trim();
}

function toCellValue_(value) {
  if (value == null) return '';
  if (Array.isArray(value)) return value.map(toCellValue_).join(' / ');
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'object') return normalizeCellText_(JSON.stringify(value));
  return normalizeCellText_(String(value));
}

function normalizeCellText_(value) {
  return String(value || '')
    .replace(/\r?\n+/g, ' / ')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 台湾雑誌シート「雑誌名」列に入力規則（マスター由来の一覧）がある場合、
 * 一覧外文字列を setValues すると書き込み全体が失敗するため、書き込み前に正規化する。
 * - マスターに照合できる → 正規の表示名（英字名など）へ
 * - mag_雑誌名を推定_ で選べる → マスター準拠の推定名のみ
 * - どちらも不可 → 雑誌名は空、候補を備考へ
 */
function mag_normalizeMagazineRowForDropdownWrite_(item) {
  if (!item || resolveSheetName_(item) !== MAGAZINE_SHEET_NAME) return item;
  var rowData = extractRowData_(item);
  if (!rowData || typeof rowData !== 'object') return item;

  var cand = String(rowData['雑誌名'] || '').trim();
  var language = mag_言語表示へ変換_(String(rowData['言語'] || '').trim() || '台湾');

  if (cand) {
    var infoHit = mag_マスターを検索_(cand, language);
    if (infoHit) {
      var canon = String(infoHit.表示英字名 || infoHit.英字名 || '').trim();
      if (canon && canon !== cand) {
        return Object.assign({}, item, {
          rowData: Object.assign({}, rowData, { '雑誌名': canon }),
        });
      }
      return item;
    }
  }

  var rawTitle = mag_firstNonEmpty_(
    rowData['原題タイトル'],
    rowData['原題商品名'],
    rowData['表紙情報'],
    rowData['表紙'],
    rowData['商品名'],
    cand
  );
  var inferred = mag_雑誌名を推定_(String(rawTitle || '').trim(), language);
  if (inferred) {
    return Object.assign({}, item, {
      rowData: Object.assign({}, rowData, { '雑誌名': inferred }),
    });
  }

  if (!cand) return item;

  var note = String(rowData['備考'] || '').trim();
  var extra = '雑誌名候補: ' + cand;
  var nextRow = Object.assign({}, rowData, {
    '雑誌名': '',
    '備考': note ? note + '\n' + extra : extra,
  });
  return Object.assign({}, item, { rowData: nextRow });
}

function mag_書込後処理_(sh, rowNum, item) {
  const rowData = extractRowData_(item);
  const map = mag_ヘッダーMap_(sh);

  let language = mag_firstNonEmpty_(
    rowData['言語'],
    map['言語'] ? sh.getRange(rowNum, map['言語']).getValue() : ''
  );
  language = mag_言語表示へ変換_(language);

  const baseTitle = mag_firstNonEmpty_(
    rowData['原題タイトル'],
    rowData['タイトル'],
    rowData['title']
  );
  const rawProductTitle = mag_firstNonEmpty_(
    rowData['原題商品名'],
    rowData['表紙情報'],
    rowData['商品名'],
    rowData['title']
  );
  const rawTitleForName = baseTitle || rawProductTitle;
  const rawTitleForParse = rawProductTitle || baseTitle;

  const incomingMagazineName = mag_firstNonEmpty_(
    rowData['雑誌名'],
    rowData['雑誌名（英字）'],
    rowData['雑誌名英字'],
    rowData['magazineName'],
    rowData['titleMagazine']
  );

  let magazineName = incomingMagazineName;
  const currentInfo = magazineName ? mag_マスターを検索_(magazineName, language) : null;
  const canonicalCurrent = currentInfo ? (currentInfo.表示英字名 || currentInfo.英字名) : '';
  if (canonicalCurrent && canonicalCurrent !== magazineName) {
    magazineName = canonicalCurrent;
  }

  const inferred = mag_雑誌名を推定_(rawTitleForName || rawTitleForParse, language);
  if (inferred && (!magazineName || magazineName !== inferred)) {
    magazineName = inferred;
  }

  if (map['言語'] && language) {
    sh.getRange(rowNum, map['言語']).setValue(language);
  }
  if (map['雑誌名'] && magazineName) {
    sh.getRange(rowNum, map['雑誌名']).setValue(magazineName);
  }
  if (map['原題タイトル'] && magazineName) {
    sh.getRange(rowNum, map['原題タイトル']).setValue(magazineName);
  }

  const workingName = magazineName || mag_雑誌名候補を整形_(rawTitleForName || rawTitleForParse);

  if (
    workingName &&
    UNKNOWN_MAGAZINE_WRITE_MODE !== 'off' &&
    !mag_雑誌マスター登録済み_(workingName, language)
  ) {
    mag_候補シートへ追加_({
      言語: language,
      雑誌名: workingName,
      原題: rawTitleForParse || rawTitleForName,
      博客來商品コード: mag_firstNonEmpty_(
        rowData['博客來商品コード'],
        rowData['サイト商品コード'],
        rowData['商品コード']
      ),
      博客來URL: mag_firstNonEmpty_(
        rowData['博客來URL'],
        rowData['リンク'],
        rowData['URL']
      ),
      rowNum,
    });
  }

  mag_行を再計算_(sh, rowNum);
}

function mag_行を再計算_(sh, row) {
  if (!sh || row < 2) return;

  const map = mag_ヘッダーMap_(sh);
  const get = name => map[name] ? sh.getRange(row, map[name]).getValue() : '';
  const set = (name, val) => { if (map[name]) sh.getRange(row, map[name]).setValue(val); };

  const languageRaw = String(get('言語') || '').trim();
  const language = mag_言語表示へ変換_(languageRaw);
  if (language && language !== languageRaw) set('言語', language);

  const baseTitle = String(get('原題タイトル') || '').trim();
  const rawProductTitle = String(get('原題商品名') || '').trim();
  const coverInfo = String(get('表紙情報') || '').trim();
  const bonusMemo = String(get('特典メモ') || '').trim();
  const rawTitleForName = baseTitle || rawProductTitle || coverInfo;
  const rawTitleForParse = rawProductTitle || coverInfo || baseTitle;

  let magazineName = String(get('雑誌名') || '').trim();
  const currentInfo = magazineName ? mag_マスターを検索_(magazineName, language) : null;
  const canonicalCurrent = currentInfo ? (currentInfo.表示英字名 || currentInfo.英字名) : '';
  if (canonicalCurrent && canonicalCurrent !== magazineName) {
    magazineName = canonicalCurrent;
    set('雑誌名', magazineName);
  }

  const inferred = mag_雑誌名を推定_(rawTitleForName, language);
  if (inferred && (!magazineName || magazineName !== inferred)) {
    magazineName = inferred;
    set('雑誌名', magazineName);
  }
  if (magazineName && baseTitle !== magazineName) {
    set('原題タイトル', magazineName);
  }

  const workingName = magazineName || mag_雑誌名候補を整形_(rawTitleForName || rawTitleForParse);
  const parsed = mag_原題を解析_(rawTitleForParse);

  let year = String(get('年') || '').trim();
  let month = String(get('月') || '').trim();
  let issue = String(get('号数') || '').trim();

  if (!year && parsed.year) { year = parsed.year; set('年', parsed.year); }
  if (!month && parsed.month) { month = parsed.month; set('月', parsed.month); }
  if (!issue && parsed.issue) { issue = parsed.issue; set('号数', parsed.issue); }

  if (!language || !workingName) return;

  const parentCode = mag_親コードを生成_(language, workingName, year, month, issue, parsed);
  if (parentCode && !parentCode.startsWith('ERROR')) set('親コード', parentCode);

  const productName = mag_商品名を生成_(language, workingName, year, month, issue, coverInfo, bonusMemo, parsed);
  if (productName) set('商品名（出品用）', productName);

  if (!get('登録日')) set('登録日', new Date());
}

function mag_ヘッダーMap_(sh) {
  const vals = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
  const map = {};
  vals.forEach((h, i) => {
    if (h) map[String(h).trim()] = i + 1;
  });
  return map;
}

function mag_マスターを検索_(magazineName, language) {
  const master = getMasterSpreadsheet_().getSheetByName(MAGAZINE_MASTER_SHEET);
  if (!master || master.getLastRow() < 2) return null;

  const map = mag_ヘッダーMap_(master);
  const rows = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getDisplayValues();
  const target = mag_正規化雑誌文字列_(magazineName);
  if (!target) return null;

  let best = null;
  let bestScore = -1;

  rows.forEach(row => {
    const info = mag_マスター行を情報化_(row, map);
    if (!info.aliases.length) return;

    info.aliases.forEach(alias => {
      const aliasNorm = mag_正規化雑誌文字列_(alias);
      if (!aliasNorm) return;

      let score = -1;
      if (target === aliasNorm) score = 300;
      else if (target.startsWith(aliasNorm)) score = 220 + aliasNorm.length;
      else if (target.includes(aliasNorm)) score = 180 + aliasNorm.length;

      if (score < 0) return;
      if (mag_言語一致_(info.languages, language)) score += 100;
      if (score > bestScore) {
        bestScore = score;
        best = info;
      }
    });
  });

  return best;
}

function mag_ルール一覧を取得_(sheetName) {
  const sh = getMasterSpreadsheet_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];

  const map = mag_ヘッダーMap_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues()
    .map(r => ({
      keyword: String(r[(map['判定キーワード'] || 1) - 1] || '').trim(),
      codeSuffix: String(r[(map['コードsuffix'] || 2) - 1] || '').trim(),
      titleDisplay: String(r[(map['タイトル表示'] || 3) - 1] || '').trim(),
      priority: Number(r[(map['優先順位'] || 4) - 1]) || 0,
      enabled: map['有効'] ? r[map['有効'] - 1] !== 'FALSE' : true,
    }))
    .filter(x => x.enabled && x.keyword)
    .sort((a, b) => b.priority - a.priority);
}

function mag_雑誌名を推定_(rawTitle, language) {
  const master = getMasterSpreadsheet_().getSheetByName(MAGAZINE_MASTER_SHEET);
  if (!master || master.getLastRow() < 2) return '';

  const map = mag_ヘッダーMap_(master);
  const rows = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getDisplayValues();
  const raw = String(rawTitle || '').trim();
  if (!raw) return '';

  let best = null;
  let bestScore = -1;

  rows.forEach(row => {
    const info = mag_マスター行を情報化_(row, map);
    if (!info.aliases.length) return;

    info.aliases.forEach(alias => {
      const score = mag_別名一致スコア_(raw, alias, info.languages, language);
      if (score > bestScore) {
        bestScore = score;
        best = info;
      }
    });
  });

  return bestScore >= 120 && best ? best.表示英字名 : '';
}function mag_雑誌マスター登録済み_(magazineName, language) {
  return !!mag_マスターを検索_(magazineName, language);
}

function mag_マスター行を情報化_(row, map) {
  const aliases = [];
  const pushIf = value => {
    const text = String(value || '').trim();
    if (text && !aliases.includes(text)) aliases.push(text);
  };

  const englishName = String(row[(map['雑誌名（英字）'] || 2) - 1] || '').trim();
  const katakanaName = String(row[(map['雑誌名（カタカナ）'] || 3) - 1] || '').trim();

  pushIf(englishName);
  Object.keys(map)
    .filter(name => /^別名\d*$/.test(name))
    .sort((a, b) => map[a] - map[b])
    .forEach(name => pushIf(row[map[name] - 1]));

  const displayEnglishName = mag_表示英字名を決定_(englishName, aliases);
  if (displayEnglishName && !aliases.includes(displayEnglishName)) aliases.push(displayEnglishName);

  return {
    略称: String(row[(map['略称コード'] || 1) - 1] || '').trim(),
    英字名: englishName,
    カタカナ名: katakanaName,
    表示英字名: displayEnglishName,
    表示カタカナ名: mag_表示カタカナ名を決定_(katakanaName),
    基本キー型: mag_基本キー型を正規化_(row[(map['基本キー型'] || 4) - 1]),
    通常タイプ結合記号: map['通常タイプ結合記号'] ? String(row[map['通常タイプ結合記号'] - 1] || '').trim() : '',
    版種ありタイプ結合記号: map['版種ありタイプ結合記号'] ? String(row[map['版種ありタイプ結合記号'] - 1] || '-').trim() : '-',
    languages: mag_対応言語一覧_(row[(map['対応言語'] || 7) - 1]),
    aliases,
  };
}

function mag_表示英字名を決定_(englishName, aliases) {
  const candidates = [];
  const seen = new Set();
  const pushCandidate = value => {
    const text = mag_英字地域表記を除去_(value);
    const key = mag_英字表示キー_(text);
    if (!text || !key || seen.has(key)) return;
    seen.add(key);
    candidates.push(text);
  };

  pushCandidate(englishName);
  (aliases || []).forEach(pushCandidate);

  if (!candidates.length) return String(englishName || '').trim();

  const latinCandidates = candidates.filter(text => /[A-Za-z]/.test(text));
  const pool = latinCandidates.length ? latinCandidates : candidates;

  let best = pool[0];
  let bestScore = mag_表示英字名スコア_(best);

  pool.forEach(candidate => {
    const score = mag_表示英字名スコア_(candidate);
    if (score > bestScore || (score === bestScore && candidate.length < best.length)) {
      best = candidate;
      bestScore = score;
    }
  });

  return mag_英字表示名を整形_(best);
}

function mag_表示英字名スコア_(value) {
  const text = String(value || '').trim();
  if (!text) return -1;

  let score = 0;
  if (/[A-Za-z]/.test(text)) score += 100;
  if (/[A-Z]/.test(text) && /[a-z]/.test(text)) score += 25;
  else if (/^[A-Z0-9 '&+./-]+$/.test(text)) score += 20;
  else if (/^[a-z0-9 '&+./-]+$/.test(text)) score += 10;
  if (!/(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/i.test(text)) score += 40;
  score += Math.max(0, 40 - text.length);
  return score;
}

function mag_英字表示名を整形_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const map = {
    'MARIE CLAIRE': 'Marie Claire',
    'VOGUE': 'VOGUE',
    'BAZAAR': 'BAZAAR',
    'GQ': 'GQ',
    'ELLE': 'ELLE',
    'ESQUIRE': 'Esquire',
    'ALLURE': 'allure',
    'MENS HEALTH': "Men's Health",
    'MEN S HEALTH': "Men's Health",
    'L OFFICIEL': "L'OFFICIEL",
    'L OFFICIEL HOMMES': "L'OFFICIEL HOMMES",
    'THE BIG ISSUE': 'THE BIG ISSUE',
    'W': 'W',
  };
  return map[mag_英字表示キー_(text)] || text;
}

function mag_英字表示キー_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function mag_英字地域表記を除去_(value) {
  return String(value || '')
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND)\b/gi, ' ')
    .replace(/\b(?:HK|TW|CN|KR|TH)\b/gi, ' ')
    .replace(/(?:韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function mag_表示カタカナ名を決定_(value) {
  return mag_カタカナ地域表記を除去_(value);
}

function mag_カタカナ地域表記を除去_(value) {
  return String(value || '')
    .trim()
    .replace(/(?:・|\s)*(?:台湾版|台湾|臺灣版|臺灣|コリア|韓国版|韓国|韓國版|韓國|香港版|香港|中国版|中国|中國版|中國|タイ版|タイ)$/u, '')
    .replace(/[・\s]+$/u, '')
    .trim();
}
function mag_対応言語一覧_(raw) {
  return String(raw || '')
    .split(',')
    .map(v => mag_言語コードへ変換_(v))
    .filter(Boolean);
}

function mag_言語一致_(supported, language) {
  if (!supported || !supported.length) return true;
  const code = mag_言語コードへ変換_(language);
  return !code || supported.includes(code);
}

function mag_別名一致スコア_(rawTitle, alias, supportedLanguages, language) {
  const raw = String(rawTitle || '').trim();
  const aliasText = String(alias || '').trim();
  if (!raw || !aliasText) return -1;

  const rawNorm = mag_正規化雑誌文字列_(raw);
  const aliasNorm = mag_正規化雑誌文字列_(aliasText);
  if (!aliasNorm) return -1;

  let score = -1;
  if (rawNorm === aliasNorm) score = 260;
  else if (rawNorm.startsWith(aliasNorm)) score = 210 + aliasNorm.length;
  else if (rawNorm.includes(aliasNorm)) score = 180 + aliasNorm.length;
  else if (aliasNorm.length <= 2 && mag_短い別名一致_(raw, aliasText)) score = 190;

  if (score < 0) return -1;
  if (mag_言語一致_(supportedLanguages, language)) score += 100;
  return score;
}

function mag_短い別名一致_(rawTitle, alias) {
  const escaped = String(alias || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(^|[^A-Z0-9])${escaped.toUpperCase()}([^A-Z0-9]|$)`);
  return re.test(String(rawTitle || '').toUpperCase());
}

function mag_正規化雑誌文字列_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[（(].*?[)）]/g, ' ')
    .replace(/&/g, ' AND ')
    .replace(/\b(KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN)\b/g, ' ')
    .replace(/(韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/g, ' ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function mag_基本キー型を正規化_(raw) {
  const s = String(raw || '').trim().toUpperCase();
  if (s === 'ISSUE' || s === 'VOL' || s === '号数型' || s === '号数') return '号数型';
  return '年月型';
}

function mag_ルール一致_(text, ruleList) {
  const hay = mag_比較用文字列_(text);
  for (let i = 0; i < ruleList.length; i += 1) {
    if (hay.includes(mag_比較用文字列_(ruleList[i].keyword))) return ruleList[i];
  }
  return { keyword: '', codeSuffix: '', titleDisplay: '', priority: 0, enabled: true };
}

function mag_原題から年月号数を抽出_(rawTitle) {
  const text = String(rawTitle || '')
    .replace(/[　]/g, ' ')
    .replace(/[／]/g, '/')
    .replace(/[－—–ー]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  let year = '';
  let month = '';
  let issue = '';
  let issueDisplay = '';

  let m = text.match(/(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月(?:號|号)?\s*\/\s*(20\d{2})/i);
  if (m) {
    month = m[1].replace(/\s*/g, '');
    year = m[2];
  }

  if (!year || !month) {
    m = text.match(/(20\d{2})\s*(?:年|\/|\-|\.)\s*(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月?(?:號|号)?/i);
    if (m) {
      year = m[1];
      month = m[2].replace(/\s*/g, '');
    }
  }

  if (!year) {
    m = text.match(/\b(20\d{2})\b/);
    if (m) year = m[1];
  }
  if (!month) {
    m = text.match(/(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月(?:號|号)?/i);
    if (m) month = m[1].replace(/\s*/g, '');
  }

  m = text.match(/\bVOL\.?\s*([0-9]{1,5})\b/i);
  if (m) {
    issue = m[1];
    issueDisplay = `Vol.${issue}`;
  }

  if (!issue) {
    m = text.match(/第\s*([0-9]{1,5})\s*(號|号|期)/i);
    if (m) {
      issue = m[1];
      issueDisplay = `第${issue}${m[2]}`;
    }
  }

  if (!issue) {
    const re = /([0-9]{1,5})\s*(號|号|期)/ig;
    let mm;
    while ((mm = re.exec(text)) !== null) {
      const prev = text.substring(Math.max(0, mm.index - 1), mm.index);
      if (prev === '月') continue;
      issue = mm[1];
      issueDisplay = `${issue}${mm[2]}`;
      break;
    }
  }

  return { year, month, issue, issueDisplay };
}

function mag_原題を解析_(rawTitle) {
  const base = mag_原題から年月号数を抽出_(rawTitle);
  return {
    year: base.year,
    month: base.month,
    issue: base.issue,
    issueDisplay: base.issueDisplay,
    edition: mag_ルール一致_(rawTitle, mag_ルール一覧を取得_(MAGAZINE_EDITION_RULE_SHEET)),
    type: mag_ルール一致_(rawTitle, mag_ルール一覧を取得_(MAGAZINE_TYPE_RULE_SHEET)),
  };
}

function mag_月コードへ変換_(monthValue) {
  const nums = String(monthValue || '').match(/\d+/g);
  if (!nums) return '';
  return nums.map(n => String(parseInt(n, 10)).padStart(2, '0')).join('');
}

function mag_月表示へ変換_(monthValue) {
  const nums = String(monthValue || '').match(/\d+/g);
  if (!nums) return '';
  const arr = nums.map(n => String(parseInt(n, 10)));
  return arr.length === 1 ? `${arr[0]}月号` : `${arr.join('・')}月号`;
}

function mag_親コードを生成_(language, magazineName, year, month, issue, parsed) {
  let info = mag_マスターを検索_(magazineName, language);

  if (!info && ALLOW_PROVISIONAL_MAGAZINE_CODE) {
    info = {
      英字名: magazineName,
      カタカナ名: '',
      略称: mag_雑誌略称候補_(magazineName),
      基本キー型: parsed.issue ? '号数型' : '年月型',
      通常タイプ結合記号: '',
      版種ありタイプ結合記号: '-',
      languages: [],
      aliases: [magazineName],
    };
  }

  if (!info || !info.略称) return 'ERROR:マスター未登録';

  let codeCore = '';
  if (info.基本キー型 === '号数型') {
    if (!issue) return 'ERROR:号数未入力';
    codeCore = `${info.略称}${issue}`;
  } else {
    if (!year || !month) return 'ERROR:年月未入力';
    const yy = String(year).slice(-2);
    const mmKey = mag_月コードへ変換_(month);
    if (!mmKey) return 'ERROR:月形式不正';
    codeCore = `${info.略称}${yy}${mmKey}`;
  }

  if (parsed.edition && parsed.edition.codeSuffix) codeCore += parsed.edition.codeSuffix;
  if (parsed.type && parsed.type.codeSuffix) {
    const joiner = parsed.edition && parsed.edition.codeSuffix
      ? (info.版種ありタイプ結合記号 || '-')
      : (info.通常タイプ結合記号 || '');
    codeCore += `${joiner}${parsed.type.codeSuffix}`;
  }

  if (!MAGAZINE_PARENT_CODE_WITH_LANGUAGE_PREFIX) return codeCore;

  const languageCode = mag_言語コードへ変換_(language);
  if (!languageCode) return 'ERROR:言語コード未対応';
  return `${languageCode}-${codeCore}`;
}

function mag_言語コードへ変換_(language) {
  const normalized = String(language || '').trim().toUpperCase();
  if (!normalized) return '';
  if (['TW', 'CN', 'HK', 'TH', 'KR', 'JP', 'US', 'EN'].includes(normalized)) return normalized;

  const map = {
    '台湾': 'TW',
    '臺灣': 'TW',
    '中国': 'CN',
    '中國': 'CN',
    '香港': 'HK',
    'タイ': 'TH',
    '泰国': 'TH',
    '泰國': 'TH',
    '韓国': 'KR',
    '韓國': 'KR',
    '日本': 'JP',
    'アメリカ': 'US',
    '英語': 'EN',
  };
  return map[String(language || '').trim()] || '';
}

function mag_言語表示へ変換_(language) {
  const code = mag_言語コードへ変換_(language);
  const map = {
    TW: '台湾',
    CN: '中国',
    HK: '香港',
    TH: 'タイ',
    KR: '韓国',
    JP: '日本',
    US: 'アメリカ',
    EN: '英語',
  };
  return map[code] || String(language || '').trim();
}

function mag_商品名を生成_(language, magazineName, year, month, issue, coverInfo, bonusMemo, parsed) {
  let info = mag_マスターを検索_(magazineName, language);

  if (!info && ALLOW_PROVISIONAL_MAGAZINE_CODE) {
    info = {
      英字名: magazineName,
      表示英字名: magazineName,
      カタカナ名: '',
      表示カタカナ名: '',
      略称: '',
      基本キー型: parsed.issue ? '号数型' : '年月型',
      通常タイプ結合記号: '',
      版種ありタイプ結合記号: '-',
      languages: [],
      aliases: [magazineName],
    };
  }

  if (!info) return '';

  const englishName = info.表示英字名 || info.英字名 || magazineName;
  const katakana = info.表示カタカナ名 || info.カタカナ名 || '';
  const baseKeyType = info.基本キー型 || '年月型';
  const languageCode = mag_言語コードへ変換_(language);
  const languageLabelMap = {
    TW: '台湾版',
    CN: '中国版',
    HK: '香港版',
    TH: 'タイ版',
  };
  const languageLabel = languageLabelMap[languageCode] || `${String(language || '').trim()}版`;

  let name = `${languageLabel} 雑誌 ${englishName}`;
  if (katakana) name += ` (${katakana})`;

  if (baseKeyType === '号数型') {
    if (!issue) return '';
    name += ` ${parsed.issueDisplay || `Vol.${issue}`}`;
    if (year && month) name += ` (${year}年${mag_月表示へ変換_(month)})`;
  } else {
    if (!year || !month) return '';
    name += ` ${year}年${mag_月表示へ変換_(month)}`;
    if (parsed.issueDisplay) name += ` ${parsed.issueDisplay}`;
  }

  if (parsed.edition && parsed.edition.titleDisplay) name += ` ${parsed.edition.titleDisplay}`;
  if (coverInfo) name += ` (${coverInfo})`;
  if (parsed.type && parsed.type.titleDisplay) name += ` ${parsed.type.titleDisplay}`;
  if (bonusMemo) name += ` ${bonusMemo}`;

  return name.replace(/\s+/g, ' ').trim();
}function mag_候補シートへ追加_(params) {
  const sh = mag_候補シートを確保_();
  const 正規化キー = `${mag_言語コードへ変換_(params.言語)}|${mag_比較用文字列_(params.雑誌名)}`;

  const 既存行 = mag_候補既存キーの行番号_(sh, 正規化キー);
  if (既存行 >= 2) {
    sh.getRange(既存行, 18).setValue(new Date());
    return;
  }

  const 言語コード = mag_言語コードへ変換_(params.言語) || '';
  const 基本キー型 = mag_雑誌キー型候補_(params.原題);
  const 備考 = [
    params.原題 ? `原題:${params.原題}` : '',
    params.博客來URL ? `URL:${params.博客來URL}` : '',
    params.博客來商品コード ? `コード:${params.博客來商品コード}` : '',
  ].filter(Boolean).join(' / ');

  const now = new Date();
  const writeRow = Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(writeRow, 1, 1, 19).setValues([[
    '未対応',
    false,
    mag_雑誌略称候補_(params.雑誌名),
    params.雑誌名 || '',
    '',
    基本キー型,
    '',
    '-',
    言語コード,
    '', '', '',
    備考,
    'books.com.tw拡張機能',
    'books_gas',
    params.rowNum || '',
    now,
    now,
    正規化キー,
  ]]);

  sh.getRange(writeRow, 2).insertCheckboxes();
}

function mag_候補既存キーの行番号_(sh, 正規化キー) {
  if (!sh || sh.getLastRow() < 2 || !正規化キー) return -1;
  const values = sh.getRange(2, 19, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i] || '').trim() === 正規化キー) return i + 2;
  }
  return -1;
}

function mag_候補シートを確保_() {
  const ss = getMasterSpreadsheet_();
  let sh = ss.getSheetByName(MAGAZINE_MASTER_CANDIDATE_SHEET);
  if (sh) return sh;

  sh = ss.insertSheet(MAGAZINE_MASTER_CANDIDATE_SHEET);

  const headers = [
    'ステータス', '反映', '略称コード候補',
    '雑誌名（英字）', '雑誌名（カタカナ）', '基本キー型候補',
    '通常タイプ結合記号', '版種ありタイプ結合記号', '対応言語',
    '別名1', '別名2', '別名3', '備考',
    '元ファイル', '元シート', '元行',
    '初回検出日時', '最終検出日時', '正規化キー'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  sh.getRange(1, 1, 1, 13)
    .setBackground('#e69138').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, 14, 1, 6)
    .setBackground('#999999').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange(2, 1, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未対応', '確認中', '登録待ち', 'マスター登録済み', '無視'], true).build()
  );

  sh.getRange(2, 2, 1000, 1).insertCheckboxes();

  sh.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['年月型', '号数型'], true).build()
  );

  const range = sh.getRange(2, 1, 1000, 13);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未対応"')
      .setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="確認中"')
      .setBackground('#fce5cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="登録待ち"')
      .setBackground('#cfe2f3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="マスター登録済み"')
      .setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="無視"')
      .setBackground('#efefef').setFontColor('#999999').setRanges([range]).build(),
  ]);

  const 列幅 = {
    1: 120, 2: 45, 3: 120, 4: 200, 5: 160,
    6: 100, 7: 75, 8: 75, 9: 75,
    10: 160, 11: 160, 12: 160, 13: 300,
    14: 160, 15: 100, 16: 60, 17: 140, 18: 140
  };
  Object.entries(列幅).forEach(([col, width]) => sh.setColumnWidth(Number(col), width));

  sh.hideColumns(19);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  return sh;
}
function mag_雑誌キー型候補_(rawTitle) {
  const text = String(rawTitle || '');
  if (/Vol\.?\s*\d+/i.test(text)) return '号数型';
  if (/第\s*\d+\s*(號|号|期)/i.test(text)) return '号数型';
  if (/\d{1,2}(?:\.\d{1,2})?\s*月(?:號|号)?\s*\/\s*20\d{2}/i.test(text)) return '年月型';
  if (/20\d{2}\s*(年|\/|\-|\.)\s*\d{1,2}\s*月/i.test(text)) return '年月型';
  return '年月型';
}

function mag_雑誌略称候補_(magazineName) {
  const cleaned = mag_正規化雑誌文字列_(magazineName).replace(/\s+/g, '');
  if (!cleaned) return '';
  return cleaned.slice(0, 8);
}

function mag_比較用文字列_(value) {
  return String(value || '')
    .trim()
    .replace(/[　\s]+/g, ' ')
    .toUpperCase();
}

function mag_雑誌名候補を整形_(rawTitle) {
  let title = String(rawTitle || '').trim();
  if (!title) return '';

  const patterns = [
    /\s+\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?\s*月(?:號|号)?\s*\/\s*20\d{2}.*$/i,
    /\s+20\d{2}\s*(?:年|\/|\-|\.)\s*\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?\s*月?(?:號|号)?.*$/i,
    /\s+VOL\.?\s*\d+.*$/i,
    /\s+NO\.?\s*\d+.*$/i,
    /\s+ISSUE\s*\d+.*$/i,
    /\s+第\s*\d+\s*(?:號|号|期).*$/i,
    /\s+\d+\s*(?:號|号|期).*$/i,
  ];

  for (let i = 0; i < patterns.length; i += 1) {
    if (patterns[i].test(title)) {
      title = title.replace(patterns[i], '');
      break;
    }
  }

  return title.replace(/\s+/g, ' ').trim();
}

function mag_firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i += 1) {
    const value = String(arguments[i] == null ? '' : arguments[i]).trim();
    if (value) return value;
  }
  return '';
}

function testExternalRequestPermission_() {
  const response = UrlFetchApp.fetch('https://www.gstatic.com/generate_204', {
    method: 'get',
    muteHttpExceptions: true,
  });

  const result = {
    ok: true,
    name: 'external_request_permission',
    statusCode: response.getResponseCode(),
    headers: response.getAllHeaders(),
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function testMangaUpdatesApi_() {
  const token = getMangaUpdatesSessionToken_();
  const response = fetchMangaUpdatesJson_('/series/search', {
    method: 'post',
    payload: { search: '全知讀者視角' },
  });

  const result = {
    ok: true,
    name: 'mangaupdates_api',
    hasToken: !!token,
    totalHits: Number((response && response.total_hits) || 0),
    firstHit: String((((response || {}).results || [])[0] || {}).hit_title || ''),
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function testTaiwanGasRuntime_() {
  const result = {
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetAccess: false,
    externalRequest: null,
    mangaupdates: null,
  };

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  result.spreadsheetAccess = !!spreadsheet;
  result.externalRequest = testExternalRequestPermission_();
  result.mangaupdates = testMangaUpdatesApi_();

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}











