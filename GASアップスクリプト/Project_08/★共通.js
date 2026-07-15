/**
 * 共通.gs【修正後完全フル版】
 * 全シート共通のユーティリティ・Works管理・SKU/タイトル生成
 */

/* ============================================================
 * ヘッダー正規化・安全列取得
 * ============================================================ */
function ヘッダー正規化_(s) {
  return String(s || '')
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/／/g, '/')
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 列番号を取得(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = ヘッダー正規化_(headers[i]);
    if (name) map[name] = i + 1;
  }
  return map;
}

function 列番号安全取得_(列マップ, 名前) {
  if (!名前) return 0;
  return 列マップ[ヘッダー正規化_(名前)] || 0;
}

function 値取得_(行データ, 列マップ, 名前) {
  const col = 列番号安全取得_(列マップ, 名前);
  if (!col) return '';
  return 正規化(行データ[col - 1]);
}

function 生値取得_(行データ, 列マップ, 名前) {
  const col = 列番号安全取得_(列マップ, 名前);
  if (!col) return '';
  return 行データ[col - 1];
}

/* ============================================================
 * マスターSpreadsheet取得（外部ファイル対応）
 * ============================================================ */
function マスターSSを取得_(cfg) {
  if (cfg && cfg.マスターファイルID) {
    return SpreadsheetApp.openById(cfg.マスターファイルID);
  }
  return SpreadsheetApp.getActive();
}

/* ============================================================
 * 基本正規化
 * ============================================================ */
function 正規化(v) {
  return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
}

function 数値変換(v) {
  const n = parseInt(String(v || '').trim(), 10);
  return Number.isFinite(n) ? n : null;
}

function 形態正規化_(v) {
  // プルダウンで選んだ形態名を勝手に変換しない。
  // 例：
  // 予約限定特装版 → 予約限定特装版 のまま
  // 限定特装版     → 限定特装版 のまま
  // 初版限定版     → 初版限定版 のまま
  return 正規化(v);
}

function キー用正規化_(v) {
  let s = 正規化(v).toLowerCase();
  s = s.replace(/^著[:：]\s*/g, '').replace(/^作[:：]\s*/g, '')
    .replace(/[［\[][^］\]]*[］\]]/g, '').replace(/[（\(][^）\)]*[）\)]/g, '').replace(/[｛\{][^｝\}]*[｝\}]/g, '')
    .replace(/[・･]/g, ' ').replace(/[～〜~]/g, '').replace(/[：:]/g, '').replace(/[、,]/g, '')
    .replace(/[。\.]/g, '').replace(/[！!]/g, '').replace(/[？?]/g, '').replace(/[『』「」]/g, '').replace(/["'"]/g, '')
    .replace(/[‐−–—―]/g, '-')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\s*[\/／]\s*/g, ' / ')
    .replace(/\s+/g, ' ')
    .replace(/-+/g, '-')
    .trim();
  return s;
}

function 原題を正規化(v) {
  let s = String(v || '').replace(/\u3000/g, ' ');
  s = s.replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[！]/g, '!').replace(/[？]/g, '?').replace(/[～〜]/g, '~');
  s = s.replace(/[：]/g, ':').replace(/[・]/g, ' ');
  s = s.replace(/[‐−–—―]/g, '-');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

/* ============================================================
 * 作者名寄せ
 * ============================================================ */
function 作者比較用正規化_(v) {
  return String(v || '')
    .replace(/\u3000/g, ' ')
    .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .toLowerCase()
    .replace(/^著[:：]\s*/g, '')
    .replace(/^作[:：]\s*/g, '')
    .replace(/[・･·•]/g, '')
    .replace(/[‐−–—―]/g, '-')
    .replace(/[()（）［］\[\]{}｛｝]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function 作者別名マップを取得_(cfg) {
  const map = {};
  const masterSS = マスターSSを取得_(cfg);
  const sh = masterSS.getSheetByName('作者名寄せマスター');
  if (!sh || sh.getLastRow() < 2) return map;

  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([別名, 正規名]) => {
    const aliasKey = 作者比較用正規化_(別名);
    const canon = 正規化(正規名);
    if (aliasKey && canon) map[aliasKey] = canon;
  });

  return map;
}

function 作者を名寄せ_(作者, 作者別名マップ) {
  const raw = 正規化(作者);
  if (!raw) return '';
  const key = 作者比較用正規化_(raw);
  return 作者別名マップ[key] || raw;
}

/* ============================================================
 * WorksKey生成
 * ============================================================ */
function WorksKeyを作る(日本語タイトル, 作者, 原題 = '', cfg = null) {
  const 正規化原題 = キー用正規化_(作品原題比較用正規化_(原題));
  const 正規化作者 = キー用正規化_(作者);
  const 正規化日本語タイトル = キー用正規化_(日本語タイトル);

  // 韓国音楽映像用：原題 + アーティスト
  if (cfg && cfg.WorksKey方式 === '原題作者') {
    const 比較用原題 = キー用正規化_(作品原題比較用正規化_(原題));
    const 比較用作者 = キー用正規化_(作品人物比較用正規化_(作者));

    if (比較用原題 && 比較用作者) {
      return '原題作者||' + 比較用原題 + '||' + 比較用作者;
    }

    if (正規化日本語タイトル && 比較用作者) {
      return '日本語作者||' + 正規化日本語タイトル + '||' + 比較用作者;
    }

    return 比較用原題 || 正規化日本語タイトル || 比較用作者 || '';
  }

  // 本・まんが系：
  // 原題がある場合でも、巻数・VOL・《》などを落とした「作品単位」の原題でWorksKeyを作る
  if (正規化原題) {
    return '原題||' + 正規化原題;
  }

  return 正規化日本語タイトル + '||' + 正規化作者;
}
/* ============================================================
 * 作品寄せ用正規化
 * - 版情報や商品ノイズを落として「同じ作品」に寄せる
 * ============================================================ */
function 作品原題比較用正規化_(v) {
  let s = 原題を正規化(v || '');
  if (!s) return '';

  s = s.toLowerCase();

  // 作品名を囲む記号
  s = s
    .replace(/[《》〈〉]/g, ' ')
    .replace(/[『』「」"'`]/g, ' ');

  // 巻数・号数・VOL表記を除去
  s = s
    .replace(/\bvol\.?\s*[0-9０-９]+(\s*[+\-~〜]\s*[0-9０-９]+)?\b/gi, ' ')
    .replace(/第\s*[0-9０-９一二三四五六七八九十百]+\s*[巻卷集冊]/g, ' ')
    .replace(/[0-9０-９]+\s*[+\-~〜]\s*[0-9０-９]+\s*[巻卷集冊]/g, ' ')
    .replace(/[0-9０-９]+\s*[巻卷集冊]/g, ' ');

  // 版・特典・商品ノイズ
  s = s
    .replace(/初版限定版|初回限定版|首刷限定版|限定版|特装版|特裝版|限定特装版|予約限定特装版/g, ' ')
    .replace(/特典付き|特典付|特典|限定特典付き|限定特典/g, ' ')
    .replace(/通常版/g, ' ');

  // 音楽映像系ノイズも一応残す
  s = s
    .replace(/\b(blu-ray|blu ray|dvd|ost|original soundtrack|lp|ep|cd|kit|kihno|platform|digipack|steelbook)\b/gi, ' ')
    .replace(/\b(4k uhd|uhd|ubd|bd)\b/gi, ' ')
    .replace(/\b(full slip|slipcase|풀슬립|스틸북|소책자|포토카드|photocard|poster|booklet|photobook|pouch|mini cd|nfc)\b/gi, ' ')
    .replace(/\b(album ver\.?|mubeat album ver\.?|limited edition|special edition|normal edition)\b/gi, ' ')
    .replace(/\b(a ver\.?|b ver\.?|c ver\.?|d ver\.?)\b/gi, ' ')
    .replace(/\b(ver\.?\s*[a-z0-9]+)\b/gi, ' ')
    .replace(/\b(version\s*[a-z0-9]+)\b/gi, ' ')
    .replace(/\b(set|box set|box|package)\b/gi, ' ');

  s = s
    .replace(/[‐−–—―]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  return s;
}

function 作品人物比較用正規化_(v) {
  let s = 作者比較用正規化_(v || '');
  if (!s) return '';

  s = s
    .replace(/variousartists/g, 'variousartists')
    .replace(/ost/g, '')
    .trim();

  return s;
}

function 作品比較キーを作る_(日本語タイトル, 作者, 原題, cfg, 言語 = '', カテゴリ = '') {
  const t = キー用正規化_(日本語タイトル || '');
  const a = 作品人物比較用正規化_(作者 || '');
  const o = 作品原題比較用正規化_(原題 || '');

  // 作品寄せは「作品名」と「作者/人物」だけで見る
  // 言語・カテゴリは比較に使わない
  const titleBase = o || t;
  if (!titleBase) return '';

  return [titleBase, a].filter(Boolean).join('||');
}

/* ============================================================
 * 既存Works候補検索
 * - WorksKey完全一致が無くても、同一作品候補を探す
 * ============================================================ */
function 既存Works候補を探す_(日本語タイトル, 作者, 原題, 言語, カテゴリ, 作品データ, cfg) {
  const result = {
    found: false,
    作品ID: '',
    worksKey: '',
    理由: ''
  };

  const target比較キー = 作品比較キーを作る_(日本語タイトル, 作者, 原題, cfg, 言語, カテゴリ);
  if (!target比較キー) return result;

  for (const [key, data] of Object.entries(作品データ.keyToData || {})) {
    const 既存比較キー = 作品比較キーを作る_(
      data.日本語タイトル || '',
      data.作者 || '',
      data.原題 || '',
      cfg,
      data.言語 || '',
      data.カテゴリ || ''
    );

    if (!既存比較キー) continue;

    if (既存比較キー === target比較キー) {
      result.found = true;
      result.作品ID = String(作品データ.keyToId[key] || '').padStart(4, '0');
      result.worksKey = key;
      result.理由 = '作品比較キー一致';
      return result;
    }
  }

  return result;
}
/* ============================================================
 * Works照合と同名疑い警告
 * ============================================================ */
function Works照合と警告(worksKey, 作者, 作品データ, 作者別名マップ = {}, cfg = null) {
  const result = { 作品ID: null, 警告: '' };

  const 既存ID = 作品データ.keyToId[worksKey];
  if (!既存ID) return result;

  result.作品ID = String(既存ID).padStart(4, '0');

  // シート設定で同名疑いチェックをOFFにしている場合は警告しない
  // 韓国音楽映像・OST・Various Artists 系の誤爆防止
  if (cfg && cfg.同名疑いチェック === false) {
    return result;
  }

  // 原題+作者方式では、作者もWorksKeyに含まれるため同名疑い警告は不要
  if (cfg && cfg.WorksKey方式 === '原題作者') {
    return result;
  }

  // 本・漫画向け：
  // 原題だけでWorksKeyを作っている場合のみ、作者違いを警告する
  if (worksKey.startsWith('原題||') && 作者) {
    const 現在作者 = 作者を名寄せ_(作者, 作者別名マップ);
    const 既存作者 = 作者を名寄せ_(
      作品データ.keyToData[worksKey]?.作者 || '',
      作者別名マップ
    );

    if (
      既存作者 &&
      作者比較用正規化_(現在作者) !== 作者比較用正規化_(既存作者)
    ) {
      result.警告 = '【同名疑い・要確認】';
    }
  }

  return result;
}

function 商品タイトル部分を抽出_(原題商品タイトル, カテゴリ = '') {
  const original = 原題を正規化(原題商品タイトル || '');
  if (!original) return '';

  let s = original
    .replace(/\bo\.?\s*s\.?\s*t\.?\b/gi, 'OST')
    .replace(/[［］]/g, ch => ch === '［' ? '[' : ']')
    .replace(/[（）]/g, ch => ch === '（' ? '(' : ')')
    .replace(/\s+/g, ' ')
    .trim();

  s = s.replace(/^\s*\[[^\]]*(?:4k|blu[-\s]?ray|블루레이|dvd|uhd|ubd)[^\]]*\]\s*/i, '');

  const dashMatch = s.match(/^(.{1,45}?)\s+-\s+(.+)$/);
  if (
    dashMatch &&
    dashMatch[1].length <= 25 &&
    !/ost/i.test(dashMatch[1]) &&
    /(?:정규|미니\s*\d+\s*집|싱글|앨범|album|no tragedy|revive|\bost\b)/i.test(dashMatch[2]) &&
    !/^(?:파우치|포토카드|아웃박스|소책자|미니\s*cd)/i.test(dashMatch[2].trim())
  ) {
    s = dashMatch[2];
  } else {
    s = s.replace(/\s+-\s+.*(?:커버|바이닐|포토|파우치|소책자|booklet|poster|photocard|album\s*ver\.?|nfc|카드|세트|아웃박스|스캐냣|스캣).*$/i, ' ');
  }

  s = s
    .replace(/^\s*(?:정규|미니|싱글|스페셜|리패키지)\s*\d+\s*(?:집|앨범)\s*/i, '')
    .replace(/^\s*\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\s*/i, '')
    .replace(/\[[^\]]*(?:180g|white\s*marble|vinyl|lp|ver\.?|한정|限定|picture|픽처|photocard|poster|cd|dvd|blu[-\s]?ray|ubd|uhd|disc)[^\]]*\]/gi, ' ')
    .replace(/\([^)]*(?:album\s*ver\.?|mubeat|pouch|파우치|mini\s*cd|미니\s*cd|photocard|포토카드|포토북|booklet|소책자|nfc|disc|cd|dvd|blu[-\s]?ray|ubd|uhd|\d+\s*종|\d+\s*p)[^)]*\)/gi, ' ')
    .replace(/\s+-\s+.*(?:커버|바이닐|포토|파우치|소책자|booklet|poster|photocard|album\s*ver\.?|nfc|카드|세트|아웃박스|스캐냣|스캣).*$/i, ' ')
    .replace(/\s*:\s*(?:스틸북|풀슬립|full\s*slip|steelbook).*$/i, ' ')
    .replace(/\b(?:4k\s*uhd|4k|uhd|ubd|2d|blu[-\s]?ray|dvd|cd|lp|vinyl|record)\b/gi, ' ')
    .replace(/(?:블루레이|스틸북|풀슬립|소책자|포토카드|포토북|북클릿|파우치|한정반|한정판|게이트폴드\s*커버|바이닐|아웃박스|스캐냣|스캣)/gi, ' ')
    .replace(/\b(?:a|b|c|d|romeo\s*cat)\s*ver\.?\b/gi, ' ')
    .replace(/\b(?:album|mubeat\s*album)\s*ver\.?\b/gi, ' ')
    .replace(/\d+\s*(?:disc|종|p)\b/gi, ' ')
    .replace(/[<＞>]+/g, ' ')
    .replace(/[『』「」"'`]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  s = s.replace(/\bost\b/gi, 'OST').trim();
  return s || original;
}

function 行の商品タイトルを取得_(行データ, 列マップ, cn, カテゴリ = '') {
  const existing = 値取得_(行データ, 列マップ, cn && cn.Worksタイトル);
  if (existing) return existing;
  const raw = 値取得_(行データ, 列マップ, cn && cn.原題商品タイトル) || 値取得_(行データ, 列マップ, cn && cn.原題);
  return 商品タイトル部分を抽出_(raw, カテゴリ);
}
/* ============================================================
 * 商品行重複チェックキー生成
 * ============================================================ */
function 商品行キーを作る(原題商品タイトル) {
  const 商品名部 = キー用正規化_(原題商品タイトル) || '';
  return 商品名部 ? `ITEM||${商品名部}` : '';
}

/* ============================================================
 * onEdit ディスパッチ
 * ============================================================ */
function メインonEdit(e) {
  console.log('--- メインonEdit START ---');

  if (!e || !e.range) {
    console.log('メインonEdit: e/rangeなし');
    return;
  }

  const sh = e.range.getSheet();
  const shName = sh.getName();
  console.log('メインonEdit: shName=' + shName);

  const ALADIN_TARGET_SHEETS = ['韓国書籍', '韓国マンガ', '韓国音楽映像'];
  const ALADIN_TRIGGERS = ['アラジンURL', 'ISBN', 'JANコード'];

  if (ALADIN_TARGET_SHEETS.includes(shName)) {
    console.log('アラジン対象シート');

    if (typeof アラジン_onEdit === 'function') {
      try {
        console.log('アラジン_onEdit 実行前');
        アラジン_onEdit(e);
        console.log('アラジン_onEdit 実行後');
      } catch (err) {
        console.log('アラジン_onEdit エラー: ' + err.message);
      }
    } else {
      console.log('アラジン_onEdit が存在しません');
    }

    const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const 編集列名 = String(ヘッダー[e.range.getColumn() - 1] || '').trim();
    console.log('メインonEdit: 編集列名=' + 編集列名);

    if (ALADIN_TRIGGERS.includes(編集列名)) {
      console.log('アラジントリガー列なので終了');
      return;
    }
  }

  if (shName === '台湾グッズ') { 台湾グッズ_onEdit(e); return; }
  if (shName === '韓国グッズ') { アラジン_onEdit(e); 韓国グッズ_onEdit(e); return; }
  if (shName === '中国価格計算') { 中国価格_onEdit(e); return; }
  if (shName === '台湾雑誌') { 台湾雑誌_onEdit(e); return; }

  const cfg = シート設定を取得(shName);
  console.log('メインonEdit: cfg=' + (cfg ? 'OK' : 'NG'));
  if (!cfg) return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  console.log('開始行=' + 開始行 + ', 行数=' + 行数);

  if (開始行 + 行数 - 1 < 2) {
    console.log('ヘッダー行なので終了');
    return;
  }

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) {
      console.log('ロック取得できず終了');
      return;
    }
  } catch (err) {
    console.log('ロックエラー: ' + err.message);
    return;
  }

  try {
    if (自己更新中か_()) {
      console.log('自己更新中フラグありで終了');
      return;
    }

    const 列マップ = 列番号を取得(sh);
    const 編集開始列 = e.range.getColumn();
    const 編集終了列 = e.range.getLastColumn();
    const 監視列番号 = (cfg.監視列 || []).map(h => 列番号安全取得_(列マップ, h)).filter(Boolean);

    console.log('編集開始列=' + 編集開始列 + ', 編集終了列=' + 編集終了列);
    console.log('監視列番号=' + JSON.stringify(監視列番号));

    const 対象列が含まれる = 監視列番号.some(c => c >= 編集開始列 && c <= 編集終了列);
    console.log('対象列が含まれる=' + 対象列が含まれる);

    if (!対象列が含まれる) {
      console.log('監視対象外なので終了');
      return;
    }

    自己更新を開始_();
    try {
      console.log('onEdit処理を実行 前');
      onEdit処理を実行(e, sh, cfg, 列マップ, 開始行, 行数);
      console.log('onEdit処理を実行 後');
    } finally {
      自己更新を終了_();
      console.log('自己更新フラグ解除');
    }
  } finally {
    lock.releaseLock();
    console.log('ロック解除');
  }
}
/* ============================================================
 * onEdit本体
 * ============================================================ */
function onEdit処理を実行(e, sh, cfg, 列マップ, 開始行, 行数) {
  const 最終列 = sh.getLastColumn();
  const cn = cfg.列名;
  const データ最終行 = データがある最終行を取得(sh, 列マップ, cn);

  const 行データ一覧 = sh.getRange(開始行, 1, 行数, 最終列).getValues();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = 作品シートを確保(ss, cfg);
  const 作品データ = 全作品データを読み込み(作品シート, cfg);

  const masterSS = マスターSSを取得_(cfg);
  const 言語マップ = 言語マップをSSから取得_(masterSS, cfg);
  const カテゴリマップ = カテゴリマップをSSから取得_(masterSS, cfg);
  const 形態マップ = 形態マップをSSから取得_(masterSS, cfg);
  const 作者別名マップ = 作者別名マップを取得_(cfg);

  // 以下そのまま

  const 出力SKU = [];
  const 出力タイトル = [];
  const 出力Worksタイトル = [];
  const 出力作品ID = [];
  const 出力SKU自動 = [];
  const 出力ステータス = [];
  const 出力作者 = [];
  const 出力形態 = [];
  const 出力重複キー = [];

  const 商品重複トラッカー = {};
  const 予約巻数マップ = {};

  const 処理対象開始行 = 開始行;
  const 処理対象終了行 = 開始行 + 行数 - 1;

  const サイト商品コード列名 =
    cn.博客來商品コード ||
    cn.サイト商品コード ||
    'サイト商品コード';

  const 取得列名直指定_ = (行データ, 列名文字列) => {
    const col = 列番号安全取得_(列マップ, 列名文字列);
    return col ? String(行データ[col - 1] || '').trim() : '';
  };

  const 商品重複キーを取得_ = (行データ) => {
    const 優先列 = Array.isArray(cfg.商品重複キー列優先順位) && cfg.商品重複キー列優先順位.length > 0
      ? cfg.商品重複キー列優先順位
      : [
          cn.ISBN,
          サイト商品コード列名,
          cn.原題商品タイトル,
          '原題商品タイトル'
        ].filter(Boolean);

    for (const 列名 of 優先列) {
      if (!列名) continue;
      const v = 取得列名直指定_(行データ, 列名);
      if (v) {
        return `${列名}||${キー用正規化_(v)}`;
      }
    }

    return '';
  };

  // 既存行の重複キーを先に記録
  if (データ最終行 >= 2) {
    const 全行データ = sh.getRange(2, 1, データ最終行 - 1, 最終列).getValues();

    for (let i = 0; i < 全行データ.length; i++) {
      const 行番号 = i + 2;
      if (行番号 >= 処理対象開始行 && 行番号 <= 処理対象終了行) continue;

      const 行データ = 全行データ[i];
      const 商品重複キー = 商品重複キーを取得_(行データ);
      if (商品重複キー) 商品重複トラッカー[商品重複キー] = true;
    }
  }

  for (let i = 0; i < 行数; i++) {
    const 行番号 = 開始行 + i;
    const 行データ = 行データ一覧[i];

    if (行番号 < 2) {
      出力SKU.push(['']);
      出力タイトル.push(['']);
      出力Worksタイトル.push(['']);
      出力作品ID.push(['']);
      出力SKU自動.push(['']);
      出力ステータス.push(['']);
      出力作者.push(['']);
      出力形態.push(['']);
      出力重複キー.push(['']);
      continue;
    }

    const 取得 = (名前) => 値取得_(行データ, 列マップ, 名前);
    const 取得生 = (名前) => 生値取得_(行データ, 列マップ, 名前);

    const 元作者 = 取得(cn.作者);
    let 作者 = 作者を名寄せ_(元作者, 作者別名マップ);

    const 言語 = 取得(cn.言語);
    const カテゴリ = 取得(cn.カテゴリ);
    const 商品タイトル = 行の商品タイトルを取得_(行データ, 列マップ, cn, カテゴリ);
const 日本語タイトル = 取得(cn.日本語タイトル) || 商品タイトル;
let 原題 = 取得(cn.原題) || 商品タイトル;
    let 形態 = 形態正規化_(取得(cn.形態));
    const 単巻数 = 取得生(cn.単巻数);
    const セット開始 = 取得生(cn.セット開始);
    const セット終了 = 取得生(cn.セット終了);
    const 特典メモ = 取得(cn.特典メモ);
const クレジット種別 = 取得(cn.クレジット種別);
    const 商品重複キー = 商品重複キーを取得_(行データ);

    if (!日本語タイトル) {
      出力SKU.push([取得生(cn.商品コード) || '']);
      出力タイトル.push([取得生(cn.タイトル) || '']);
      出力Worksタイトル.push([商品タイトル || 取得生(cn.Worksタイトル) || '']);
      出力作品ID.push([取得生(cn.作品ID) || '']);
      出力SKU自動.push([取得生(cn.SKU自動) || '']);
      出力ステータス.push([取得生(cn.コードステータス) || '']);
      出力作者.push([作者 || 取得生(cn.作者) || '']);
      出力形態.push([形態 || 取得生(cn.形態) || '']);
      出力重複キー.push([商品重複キー || 取得生(cn.重複チェックキー) || '']);
      continue;
    }

    const 全条件揃い = !!(日本語タイトル && 作者 && 言語 && カテゴリ);
    let 作品ID = '';
    let 同名警告 = '';
    let worksKey = '';

    // 予約中にすでに作品IDがある場合は、それを優先する
    const 既存作品ID = String(取得生(cn.作品ID) || '').trim();

    if (作者) {
      worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題, cfg);

      // 予約中は既存作品IDを優先して、Worksを増やさない
      if (
        cfg.予約時既存作品ID優先 === true &&
        既存作品ID &&
        既存作品ID !== '????'
      ) {
        作品ID = String(既存作品ID).padStart(4, '0');

        if (!作品データ.予約上書きRows) 作品データ.予約上書きRows = [];

        作品データ.予約上書きRows.push({
          作品ID,
          worksKey,
          日本語タイトル,
          作者,
          原題: 原題を正規化(原題)
        });

        作品データ.keyToId[worksKey] = 作品ID;
        作品データ.keyToData[worksKey] = {
  日本語タイトル,
  作者,
  原題
};

      } else {
        const 照合結果 = Works照合と警告(
  worksKey,
  作者,
  作品データ,
  作者別名マップ,
  cfg
);

               if (照合結果.作品ID) {
          作品ID = 照合結果.作品ID;
          同名警告 = 照合結果.警告;

          const 既存 = 作品データ.keyToData[worksKey];
          if (既存) {
            if (既存.作者) 作者 = 作者を名寄せ_(既存.作者, 作者別名マップ);
            if (!原題 && 既存.原題) 原題 = 既存.原題;
          }

             } else if (全条件揃い) {
          const 候補 = 既存Works候補を探す_(
            日本語タイトル,
            作者,
            原題,
            言語,
            カテゴリ,
            作品データ,
            cfg
          );

          if (候補.found && 候補.作品ID) {
            作品ID = 候補.作品ID;
            同名警告 = '【既存Works候補に統一】';

            const 既存 = 作品データ.keyToData[候補.worksKey];
            if (既存) {
              if (既存.作者) 作者 = 作者を名寄せ_(既存.作者, 作者別名マップ);
              if (!原題 && 既存.原題) 原題 = 既存.原題;
            }

            作品データ.keyToId[worksKey] = 作品ID;
            作品データ.keyToData[worksKey] = {
              日本語タイトル,
              作者,
              原題
            };
          } else {
            作品データ.maxId++;
            作品ID = String(作品データ.maxId).padStart(4, '0');
            // 発行した瞬間にハイウォーターを永続化（行削除後の同ID再発行を根絶）
            採番ハイウォーター更新(cfg, 作品データ.maxId);

            作品データ.keyToId[worksKey] = 作品ID;
            作品データ.keyToRow[worksKey] = null;
            作品データ.keyToData[worksKey] = {
              日本語タイトル,
              作者,
              原題
            };

            作品データ.newRows.push([
              worksKey,
              作品ID,
              日本語タイトル,
              作者,
              原題を正規化(原題),
              '',
              '',
              '',
              '',
              ''
            ]);
          }

        } else {
          作品ID = '????';
        }
      }

    } else {
      作品ID = '????';
    }

    let 巻数警告 = '';

    if (商品重複キー) {
      if (商品重複トラッカー[商品重複キー]) {
        巻数警告 = '【重複注意】';
      }
      商品重複トラッカー[商品重複キー] = true;
    }

    const 巻数 = 数値変換(単巻数);
    const セット終了巻 = 数値変換(セット終了);
    const 最新巻候補 = 巻数 != null ? 巻数 : セット終了巻;

    if (最新巻候補 != null && 作品ID !== '????' && worksKey) {
      if (!予約巻数マップ[worksKey]) 予約巻数マップ[worksKey] = [];
      予約巻数マップ[worksKey].push(最新巻候補);
    }

    const 言語コード = 言語 ? (言語マップ[言語] || 'XX') : '';
    const カテゴリコード = カテゴリ ? (カテゴリマップ[カテゴリ] || 'XX') : '';
    const 形態コード = 形態コードを取得_(形態, 形態マップ);

    const SKU = SKUを段階生成(
      言語コード,
      形態コード,
      作品ID,
      カテゴリコード,
      単巻数,
      セット開始,
      セット終了
    );

    const タイトル =
  同名警告 +
  巻数警告 +
  タイトルを段階生成(
    cfg,
    言語,
    カテゴリ,
    形態,
    日本語タイトル,
    単巻数,
    セット開始,
    セット終了,
    作者,
    原題,
    特典メモ,
    形態マップ,
    クレジット種別
  );

    let ステータス = '';
    if (同名警告) ステータス = 同名警告;
    else if (巻数警告) ステータス = '【重複注意】';
    else if (全条件揃い) ステータス = '商品コード(予約)';
    else ステータス = '入力中...';

    出力SKU.push([SKU]);
    出力タイトル.push([タイトル]);
    出力Worksタイトル.push([商品タイトル || '']);
    出力作品ID.push([作品ID]);
    出力SKU自動.push([SKU]);
    出力ステータス.push([ステータス]);
    出力作者.push([作者 || 取得生(cn.作者) || '']);
    出力形態.push([形態 || 取得生(cn.形態) || '']);
    出力重複キー.push([商品重複キー || '']);
  }

  const col商品コード = 列番号安全取得_(列マップ, cn.商品コード);
  const colタイトル = 列番号安全取得_(列マップ, cn.タイトル);
  const colWorksタイトル = 列番号安全取得_(列マップ, cn.Worksタイトル);
  const col作品ID = 列番号安全取得_(列マップ, cn.作品ID);
  const colSKU自動 = 列番号安全取得_(列マップ, cn.SKU自動);
  const colコードステータス = 列番号安全取得_(列マップ, cn.コードステータス);
  const col作者 = 列番号安全取得_(列マップ, cn.作者);
  const col形態 = 列番号安全取得_(列マップ, cn.形態);
  const col重複チェックキー = 列番号安全取得_(
    列マップ,
    cn.重複チェックキー || cfg.商品重複キー出力列
  );

  if (col商品コード) sh.getRange(開始行, col商品コード, 行数, 1).setValues(出力SKU);
  if (colタイトル) sh.getRange(開始行, colタイトル, 行数, 1).setValues(出力タイトル);
  if (colWorksタイトル) sh.getRange(開始行, colWorksタイトル, 行数, 1).setValues(出力Worksタイトル);
  if (col作品ID) sh.getRange(開始行, col作品ID, 行数, 1).setValues(出力作品ID);
  if (colSKU自動) sh.getRange(開始行, colSKU自動, 行数, 1).setValues(出力SKU自動);
  if (colコードステータス) sh.getRange(開始行, colコードステータス, 行数, 1).setValues(出力ステータス);
  if (col作者) sh.getRange(開始行, col作者, 行数, 1).setValues(出力作者);
  if (col形態) sh.getRange(開始行, col形態, 行数, 1).setValues(出力形態);
  if (col重複チェックキー) sh.getRange(開始行, col重複チェックキー, 行数, 1).setValues(出力重複キー);

  作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ, cfg);
}

/* ============================================================
 * Works更新（onEdit専用）
 * ============================================================ */
function 作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ, cfg) {
  // 既存のWorksKey補正
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    const 最終行 = 作品シート.getLastRow();
    if (最終行 >= 2) {
      const 全Works = 作品シート.getRange(2, 1, 最終行 - 1, 1).getValues();

      for (const upd of 作品データ.keyUpdates) {
        全Works[upd.行 - 2][0] = upd.key;
      }

      作品シート.getRange(2, 1, 最終行 - 1, 1).setValues(全Works);
    }

    作品データ.keyUpdates = [];
  }

  // 新規Works追加
  if (作品データ.newRows && 作品データ.newRows.length > 0) {
    const 開始行 = Math.max(2, 作品シート.getLastRow() + 1);
    作品シート
      .getRange(開始行, 1, 作品データ.newRows.length, cfg.作品列数)
      .setValues(作品データ.newRows);

    for (let i = 0; i < 作品データ.newRows.length; i++) {
      const key = String(作品データ.newRows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 開始行 + i;
    }
  }

  // 予約中の既存作品ID優先によるWorks上書き
  if (作品データ.予約上書きRows && 作品データ.予約上書きRows.length > 0) {
  const 最終行 = 作品シート.getLastRow();
  let 全Works = [];
  const idToIndex = {};
  let 更新あり = false;
  const 追加Rows = [];

  if (最終行 >= 2) {
    全Works = 作品シート
      .getRange(2, 1, 最終行 - 1, cfg.作品列数)
      .getValues();

    for (let i = 0; i < 全Works.length; i++) {
      const id = String(全Works[i][1] || '').trim().padStart(4, '0');
      if (id) idToIndex[id] = i;
    }
  }

  for (const row of 作品データ.予約上書きRows) {
    const 作品ID = String(row.作品ID || '').trim().padStart(4, '0');
    if (!作品ID || 作品ID === '????') continue;

    const 新WorksRow = [
      row.worksKey || '',
      作品ID,
      row.日本語タイトル || '',
      row.作者 || '',
      row.原題 || '',
      '',
      '',
      '',
      '',
      ''
    ];

    if (idToIndex[作品ID] != null) {
      const idx = idToIndex[作品ID];

      全Works[idx][0] = 新WorksRow[0];
      全Works[idx][1] = 新WorksRow[1];
      全Works[idx][2] = 新WorksRow[2];
      全Works[idx][3] = 新WorksRow[3];
      全Works[idx][4] = 新WorksRow[4];

      作品データ.keyToRow[row.worksKey] = idx + 2;
      更新あり = true;

    } else {
      追加Rows.push(新WorksRow);
    }
  }

  if (更新あり && 全Works.length > 0) {
    作品シート
      .getRange(2, 1, 全Works.length, cfg.作品列数)
      .setValues(全Works);
  }

  if (追加Rows.length > 0) {
    const 追加開始行 = Math.max(2, 作品シート.getLastRow() + 1);

    作品シート
      .getRange(追加開始行, 1, 追加Rows.length, cfg.作品列数)
      .setValues(追加Rows);

    for (let i = 0; i < 追加Rows.length; i++) {
      const key = String(追加Rows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 追加開始行 + i;
    }
  }

  作品データ.予約上書きRows = [];
}

  // 予約込み最新巻・最新数の更新
  const 最終行 = 作品シート.getLastRow();

  if (最終行 >= 2 && Object.keys(予約巻数マップ).length > 0) {
    const IJ列 = 作品シート.getRange(2, 9, 最終行 - 1, 2).getValues();
    const now = new Date();

    for (const [key, 巻数リスト] of Object.entries(予約巻数マップ)) {
      const 行番号 = 作品データ.keyToRow[key];
      if (!行番号) continue;

      const 今回最大 = Math.max(...巻数リスト);
      const 既存 = 作品データ.keyTo予約最新巻[key] || 0;

      IJ列[行番号 - 2][0] = Math.max(今回最大, 既存);
      IJ列[行番号 - 2][1] = now;
    }

    作品シート.getRange(2, 9, 最終行 - 1, 2).setValues(IJ列);
  }

  作品データ.newRows = [];
}

/* ============================================================
 * Works更新（確定版）
 * ============================================================ */
function 作品データを更新_確定(作品シート, 作品データ, cfg) {
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    const 最終行 = 作品シート.getLastRow();
    if (最終行 >= 2) {
      const 全Works = 作品シート.getRange(2, 1, 最終行 - 1, 1).getValues();
      for (const upd of 作品データ.keyUpdates) {
        全Works[upd.行 - 2][0] = upd.key;
      }
      作品シート.getRange(2, 1, 最終行 - 1, 1).setValues(全Works);
    }
    作品データ.keyUpdates = [];
  }

  if (作品データ.newRows.length > 0) {
    const 開始行 = Math.max(2, 作品シート.getLastRow() + 1);
    作品シート.getRange(開始行, 1, 作品データ.newRows.length, cfg.作品列数).setValues(作品データ.newRows);
    for (let i = 0; i < 作品データ.newRows.length; i++) {
      const key = String(作品データ.newRows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 開始行 + i;
    }
  }

  const 最終行 = 作品シート.getLastRow();
  if (最終行 < 2) return;

  const 更新対象キー = Object.entries(作品データ.keyToVols).filter(([_, vols]) => vols && vols.size > 0);
  if (更新対象キー.length === 0 && (!作品データ.keyUpdates || 作品データ.keyUpdates.length === 0)) return;

  const 全Works = 作品シート.getRange(2, 1, 最終行 - 1, cfg.作品列数).getValues();
  const now = new Date();
  let 更新あり = false;

  for (const [key, vols] of 更新対象キー) {
    const 行番号 = 作品データ.keyToRow[key];
    if (!行番号) continue;
    const idx = 行番号 - 2;
    if (idx < 0 || idx >= 全Works.length) continue;

    const arr = Array.from(vols).sort((a, b) => a - b);
    全Works[idx][5] = arr.join(',');
    全Works[idx][6] = Math.max(...arr);
    全Works[idx][7] = now;
    更新あり = true;
  }

  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    for (const upd of 作品データ.keyUpdates) {
      const idx = upd.行 - 2;
      if (idx >= 0 && idx < 全Works.length) 全Works[idx][0] = upd.key;
    }
    更新あり = true;
    作品データ.keyUpdates = [];
  }

  if (更新あり) 作品シート.getRange(2, 1, 全Works.length, cfg.作品列数).setValues(全Works);
  作品データ.newRows = [];
}

/* ============================================================
 * WorksKey再正規化
 * ============================================================ */
function WorksKey再正規化を実行(作品シート, cfg) {
  const 作者別名マップ = 作者別名マップを取得_(cfg);
  if (!作品シート || 作品シート.getLastRow() < 2) return { キー更新数: 0, 統合数: 0, 削除行数: 0 };

  const 最終行 = 作品シート.getLastRow();
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, cfg.作品列数).getValues();

  const タイトル作者グループ = new Map();
  for (let i = 0; i < データ.length; i++) {
    const t = 正規化(データ[i][2] || '');
    const a = 作者を名寄せ_(データ[i][3] || '', 作者別名マップ);
    if (!t || !a) continue;

    const groupKey = キー用正規化_(t) + '||' + キー用正規化_(a);
    if (!タイトル作者グループ.has(groupKey)) タイトル作者グループ.set(groupKey, []);
    タイトル作者グループ.get(groupKey).push(i);
  }

  for (const [_, 行インデックス] of タイトル作者グループ.entries()) {
    if (行インデックス.length < 2) continue;
    const 原題あり = 行インデックス.find(i => 正規化(データ[i][4] || ''));
    if (原題あり == null) continue;

    const 正規化原題 = 原題を正規化(データ[原題あり][4]);
    for (const i of 行インデックス) {
      if (i !== 原題あり && !正規化(データ[i][4] || '')) データ[i][4] = 正規化原題;
    }
  }

  const keyMap = new Map();
  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const t = 正規化(r[2] || '');
    const a = 作者を名寄せ_(r[3] || '', 作者別名マップ);
    const o = 正規化(r[4] || '');
    if (!t && !a) continue;

    const key = WorksKeyを作る(t, a, o, cfg);
    const 元key = String(r[0] || '').trim();

    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: r, 元key, キー変更: 元key !== key });
  }

  let キー更新数 = 0;
  let 統合数 = 0;
  const 削除行 = [];
  const 更新データ = データ.map(r => r.slice());

  for (const [key, 行配列] of keyMap.entries()) {
    if (行配列.length === 1) {
      const idx = 行配列[0].行 - 2;
      const 正規作者 = 作者を名寄せ_(更新データ[idx][3] || '', 作者別名マップ);

      if (更新データ[idx][3] !== 正規作者) {
        更新データ[idx][3] = 正規作者;
      }

      if (行配列[0].キー変更) {
        更新データ[idx][0] = key;
        キー更新数++;
      }
    } else {
      行配列.sort((a, b) => {
        const a原題 = 正規化(a.データ[4] || '') ? 0 : 1;
        const b原題 = 正規化(b.データ[4] || '') ? 0 : 1;
        if (a原題 !== b原題) return a原題 - b原題;
        return b.行 - a.行;
      });

      const 残す = 行配列[0];
      更新データ[残す.行 - 2][3] = 作者を名寄せ_(更新データ[残す.行 - 2][3] || '', 作者別名マップ);
      更新データ[残す.行 - 2][0] = key;

      const 全巻 = new Set();
      for (const item of 行配列) {
        String(item.データ[5] || '').split(',').forEach(v => {
          const n = parseInt(String(v).trim(), 10);
          if (!isNaN(n)) 全巻.add(n);
        });
      }

      if (全巻.size > 0) {
        const arr = Array.from(全巻).sort((a, b) => a - b);
        更新データ[残す.行 - 2][5] = arr.join(',');
        更新データ[残す.行 - 2][6] = Math.max(...arr);
        更新データ[残す.行 - 2][7] = new Date();
      }

      let 最大予約巻 = 0;
      for (const item of 行配列) {
        const v = parseInt(String(item.データ[8] || '0'), 10);
        if (!isNaN(v) && v > 最大予約巻) 最大予約巻 = v;
      }

      if (最大予約巻 > 0) {
        更新データ[残す.行 - 2][8] = 最大予約巻;
        更新データ[残す.行 - 2][9] = new Date();
      }

      if (!正規化(残す.データ[4] || '')) {
        for (let j = 1; j < 行配列.length; j++) {
          const t = 正規化(行配列[j].データ[4] || '');
          if (t) {
            更新データ[残す.行 - 2][4] = t;
            break;
          }
        }
      }

      for (let j = 1; j < 行配列.length; j++) {
        削除行.push(行配列[j].行);
      }
      統合数++;
    }
  }

  作品シート.getRange(2, 1, 更新データ.length, cfg.作品列数).setValues(更新データ);
  行を一括削除(作品シート, 削除行);

  return { キー更新数, 統合数, 削除行数: 削除行.length };
}

/* ============================================================
 * WorksID振り直し
 * ============================================================ */
function WorksID振り直しを実行(作品シート, cfg) {
  // 【封印 2026-07-15】作品IDは永久ID。1から連番に詰め直すと、登録済み商品行や
  // Yahoo等の外部に出たコードは古いIDのまま残り、同じIDが別作品を指す
  // 「作品IDずれ」になる（一括更新が毎回これを呼んでいたのが再発源だった）。
  // 現在の採番は max+1+ハイウォーター方式のため、欠番はそのままで問題ない。
  // 振り直しがどうしても必要な場合は、商品行・外部コードの一括書き換えと
  // セットで別途設計すること。
  return { 変更数: 0, 旧新マップ: {} };
}

/* ============================================================
 * Works重複検出
 * ============================================================ */
function Works重複を検出(作品シート, cfg) {
  if (!作品シート || 作品シート.getLastRow() < 2) return [];

  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, cfg.作品列数).getValues();
  const keyMap = new Map();

  for (let i = 0; i < データ.length; i++) {
    const t = 正規化(データ[i][2]);
    const a = 正規化(データ[i][3]);
    const o = 正規化(データ[i][4]);
    if (!t || !a) continue;

    const key = WorksKeyを作る(t, a, o, cfg);
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: データ[i] });
  }

  return Array.from(keyMap.entries())
    .filter(([_, v]) => v.length > 1)
    .map(([key, 行配列]) => ({ key, 行配列 }));
}

/* ============================================================
 * Works孤立エントリー削除
 * ============================================================ */
function Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, cfg) {
  const result = { 削除数: 0, 削除リスト: [] };
  if (!作品シート || 作品シート.getLastRow() < 2) return result;

  const cn = cfg.列名;
  const 最終行 = データがある最終行を取得(マスターシート, 列マップ, cn);
  const 使用中Keys = new Set();

  if (最終行 >= 2) {
    const 全データ = マスターシート.getRange(2, 1, 最終行 - 1, マスターシート.getLastColumn()).getValues();
    for (const r of 全データ) {
      const t = 値取得_(r, 列マップ, cn.日本語タイトル) || 値取得_(r, 列マップ, cn.Worksタイトル);
      const a = 値取得_(r, 列マップ, cn.作者);
      const o = 値取得_(r, 列マップ, cn.原題);
      if (t && a) {
        使用中Keys.add(WorksKeyを作る(t, a, o, cfg));
        使用中Keys.add(キー用正規化_(t) + '||' + キー用正規化_(a));
      }
    }
  }

  const worksData = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, cfg.作品列数).getValues();
  const 削除行 = [];

  for (let i = 0; i < worksData.length; i++) {
    const t = 正規化(worksData[i][2] || '');
    const a = 正規化(worksData[i][3] || '');
    const o = 正規化(worksData[i][4] || '');

    if (!t && !a) {
      削除行.push(i + 2);
      continue;
    }

    const computedKey = WorksKeyを作る(t, a, o, cfg);
    const storedKey = String(worksData[i][0] || '').trim();
    const oldKey = キー用正規化_(t) + '||' + キー用正規化_(a);

    if (!使用中Keys.has(computedKey) && !使用中Keys.has(storedKey) && !使用中Keys.has(oldKey)) {
      削除行.push(i + 2);
      result.削除リスト.push(`ID:${worksData[i][1]} ${worksData[i][2]} / ${worksData[i][3]}`);
    }
  }

  行を一括削除(作品シート, 削除行);
  result.削除数 = 削除行.length;
  return result;
}

/* ============================================================
 * 作品シート確保・全データ読み込み
 * ============================================================ */
function 作品シートを確保(ss, cfg) {
  let sh = ss.getSheetByName(cfg.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(cfg.作品シート名);
    sh.getRange(1, 1, 1, cfg.作品列数).setValues([cfg.作品ヘッダー]);
  } else {
    const 現在列数 = sh.getLastColumn();
    if (現在列数 < cfg.作品列数) {
      for (let i = 現在列数; i < cfg.作品列数; i++) {
        sh.getRange(1, i + 1).setValue(cfg.作品ヘッダー[i]);
      }
    }
  }
  return sh;
}

/* ============================================================
 * 作品ID採番ハイウォーターマーク
 *
 * 従来の採番は「Works走査のmaxId + 1」だけで、Works行を削除すると
 * 同じIDが別作品に再発行される欠陥があった（台湾側で実際に発生した事故と同型）。
 * DocumentProperties と スプレッドシート上の「採番管理」シート（Worksシート名ごと
 * に1行）に発行済み最大値を永続化し、走査maxと合成する。欠番は永久欠番。
 * セル側は複数プロジェクト（例: グッズ=Project_06 と 書籍系=Project_05）から
 * 見える共有ストア、props側はセルが壊れた場合の縮退用。
 * ============================================================ */
function 採番ハイウォーターセル_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('採番管理');
  if (!sh) {
    try {
      sh = ss.insertSheet('採番管理');
      sh.getRange(1, 1).setValue('Worksシート名（削除・編集禁止。作品ID採番のハイウォーター）');
      sh.getRange(1, 2).setValue('発行済み最大値');
      sh.hideSheet();
    } catch (e) {
      sh = ss.getSheetByName('採番管理');
    }
  }
  if (!sh) return null;
  const 最終行 = sh.getLastRow();
  if (最終行 >= 2) {
    const 名前一覧 = sh.getRange(2, 1, 最終行 - 1, 1).getDisplayValues();
    for (let i = 0; i < 名前一覧.length; i++) {
      if (String(名前一覧[i][0]).trim() === cfg.作品シート名) return sh.getRange(i + 2, 2);
    }
  }
  const 行 = Math.max(最終行 + 1, 2);
  sh.getRange(行, 1).setValue(cfg.作品シート名);
  return sh.getRange(行, 2);
}

function 採番ハイウォーター読取(cfg) {
  let props値 = 0;
  try {
    props値 = parseInt(
      PropertiesService.getDocumentProperties()
        .getProperty('作品ID_ハイウォーター_' + cfg.作品シート名) || '0', 10) || 0;
  } catch (e) {}
  let セル値 = 0;
  try {
    const range = 採番ハイウォーターセル_(cfg);
    if (range) セル値 = parseInt(String(range.getDisplayValue()).replace(/\D/g, ''), 10) || 0;
  } catch (e) {}
  return Math.max(props値, セル値);
}

function 採番ハイウォーター更新(cfg, 値) {
  try {
    PropertiesService.getDocumentProperties()
      .setProperty('作品ID_ハイウォーター_' + cfg.作品シート名, String(値));
  } catch (e) {}
  try {
    const range = 採番ハイウォーターセル_(cfg);
    if (range) {
      range.setValue(値);
      SpreadsheetApp.flush();
    }
  } catch (e) {}
}

function 全作品データを読み込み(作品シート, cfg) {
  const result = {
    keyToId: {},
    keyToData: {},
    keyToRow: {},
    keyToVols: {},
    keyTo予約最新巻: {},
    maxId: 0,
    newRows: [],
    keyUpdates: [],
    titleToKey: {}
  };

  // 発行済みハイウォーターを合成（Works行が削除・全消去されてもIDを再発行しない）
  const 保存最大 = 採番ハイウォーター読取(cfg);
  if (保存最大 > result.maxId) result.maxId = 保存最大;

  const 最終行 = 作品シート.getLastRow();
  if (最終行 < 2) return result;

  const 列数 = Math.max(作品シート.getLastColumn(), cfg.作品列数);
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, Math.min(列数, cfg.作品列数)).getValues();

  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const idStr = String(r[1] == null ? '' : r[1]).trim();
    if (!idStr) continue;

    const t = 正規化(r[2] || '');
    const a = 正規化(r[3] || '');
    const o = 正規化(r[4] || '');
    let key = (t && a) ? WorksKeyを作る(t, a, o, cfg) : String(r[0] || '').trim();
    if (!key) continue;

    const 保存済みKey = String(r[0] || '').trim();
    if (保存済みKey !== key) result.keyUpdates.push({ 行: i + 2, key });

    if (result.keyToId[key]) {
      const 既存ID = parseInt(result.keyToId[key], 10);
      const 新ID = parseInt(idStr, 10);
      if (!isNaN(新ID) && !isNaN(既存ID) && 新ID >= 既存ID) {
        String(r[5] || '').split(',').forEach(v => {
          const n = parseInt(String(v).trim(), 10);
          if (!isNaN(n)) {
            if (!result.keyToVols[key]) result.keyToVols[key] = new Set();
            result.keyToVols[key].add(n);
          }
        });
        const num = parseInt(idStr, 10);
        if (!isNaN(num) && num > result.maxId) result.maxId = num;
        continue;
      }
    }

    result.keyToId[key] = idStr.padStart(4, '0');
    result.keyToRow[key] = i + 2;
    result.keyToData[key] = {
  日本語タイトル: 正規化(r[2] || ''),
  作者: 正規化(r[3] || ''),
  原題: 正規化(r[4] || ''),
  言語: '',
  カテゴリ: ''
};
    if (!result.keyToVols[key]) result.keyToVols[key] = new Set();

    String(r[5] || '').split(',').forEach(v => {
      const n = parseInt(String(v).trim(), 10);
      if (!isNaN(n)) result.keyToVols[key].add(n);
    });

    const 予約巻 = parseInt(String(r[8] || '0'), 10);
    if (!isNaN(予約巻) && 予約巻 > 0) result.keyTo予約最新巻[key] = 予約巻;

    const num = parseInt(idStr, 10);
    if (!isNaN(num) && num > result.maxId) result.maxId = num;

    if (t && !key.startsWith('原題||')) {
      const titleKey = キー用正規化_(t);
      if (!result.titleToKey[titleKey]) result.titleToKey[titleKey] = key;
    }
  }

  return result;
}

/* ============================================================
 * ① 確定発行（修正版）
 * - 日本語タイトルだけの fallback を使わない
 * - Works寄せは原題タイトルベース
 * ============================================================ */
function 確定発行を実行(cfg) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.マスターシート名);
  if (!sh) return;

  const cn = cfg.列名;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ, cn);
  if (最終行 < 2) return;

  const masterSS = マスターSSを取得_(cfg);
  const 言語マップ = 言語マップをSSから取得_(masterSS, cfg);
  const カテゴリマップ = カテゴリマップをSSから取得_(masterSS, cfg);
  const 形態マップ = 形態マップをSSから取得_(masterSS, cfg);
  const 作者別名マップ = 作者別名マップを取得_(cfg);

  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターが見つかりません。マスターファイルIDを確認してください。');
    return;
  }

  const totalCols = sh.getLastColumn();
  const 全データ = sh.getRange(2, 1, 最終行 - 1, totalCols).getValues();
  const 作品シート = 作品シートを確保(ss, cfg);

  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);

  try {
    // 高速化: 必要な時だけ Works 再正規化
    if (cfg.確定発行前にWorks再正規化 === true) {
      WorksKey再正規化を実行(作品シート, cfg);
    }

    const 作品データ = 全作品データを読み込み(作品シート, cfg);

    const colWorksタイトル = 列番号安全取得_(列マップ, cn.Worksタイトル);
    const col作品ID = 列番号安全取得_(列マップ, cn.作品ID);
    const colSKU自動 = 列番号安全取得_(列マップ, cn.SKU自動);
    const col商品コード = 列番号安全取得_(列マップ, cn.商品コード);
    const colタイトル = 列番号安全取得_(列マップ, cn.タイトル);
    const colステータス = 列番号安全取得_(列マップ, cn.コードステータス);
    const col作者 = 列番号安全取得_(列マップ, cn.作者);
    const col登録状況 = 列番号安全取得_(列マップ, cn.登録状況);
    const col発行チェック = 列番号安全取得_(列マップ, cn.発行チェック);

    let 発行数 = 0;

    const 変更map作品ID = {};
    const 変更mapSKU自動 = {};
    const 変更map商品コード = {};
    const 変更mapタイトル = {};
    const 変更mapWorksタイトル = {};
    const 変更mapステータス = {};
    const 変更map作者 = {};
    const 変更map登録状況 = {};

    // 既存作品ID → WorksKey を先に作っておく
    const 作品IDから既存WorksKey = {};
    for (const [k, id] of Object.entries(作品データ.keyToId || {})) {
      const 正規ID = String(id || '').padStart(4, '0');
      if (正規ID) {
        作品IDから既存WorksKey[正規ID] = k;
      }
    }

    function 列をまとめて書き込む(colNum, valueMap) {
      if (!colNum) return;

      const rows = Object.keys(valueMap).map(Number).sort((a, b) => a - b);
      if (rows.length === 0) return;

      let start = rows[0];
      let buf = [[valueMap[rows[0]]]];

      for (let k = 1; k <= rows.length; k++) {
        if (k < rows.length && rows[k] === rows[k - 1] + 1) {
          buf.push([valueMap[rows[k]]]);
        } else {
          sh.getRange(start, colNum, buf.length, 1).setValues(buf);
          if (k < rows.length) {
            start = rows[k];
            buf = [[valueMap[rows[k]]]];
          }
        }
      }
    }

    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const sheetRow = i + 2;

      const 取得 = (名前) => 値取得_(r, 列マップ, 名前);
      const 取得生 = (名前) => 生値取得_(r, 列マップ, 名前);

      if (取得生(cn.発行チェック) !== true) continue;

      const 言語 = 取得(cn.言語);
      const カテゴリ = 取得(cn.カテゴリ);
      const 商品タイトル = 行の商品タイトルを取得_(r, 列マップ, cn, カテゴリ);
const 日本語タイトル = 取得(cn.日本語タイトル) || 商品タイトル;
let 作者 = 作者を名寄せ_(取得(cn.作者), 作者別名マップ);
let 原題 = 取得(cn.原題) || 商品タイトル;
      const 形態 = 形態正規化_(取得(cn.形態));
      const クレジット種別 = cn.クレジット種別 ? 取得(cn.クレジット種別) : '';

      if (!日本語タイトル || !作者 || !言語 || !カテゴリ) {
        if (colステータス) 変更mapステータス[sheetRow] = '入力中...';
        continue;
      }

      const 言語コード = 言語マップ[言語] || 'XX';
      const カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const 形態コード = 形態コードを取得_(形態, 形態マップ);

      const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題, cfg);
      let 更新先WorksKey = worksKey;

      let 作品ID = 作品データ.keyToId[worksKey] || '';
      const 既存 = 作品データ.keyToData[worksKey];

      const 現在作品ID = 正規化(取得(cn.作品ID));
      const 現在商品コード = 正規化(取得(cn.商品コード));
      const 現在SKU = 正規化(取得(cn.SKU自動));

      const 行既存作品ID候補 =
  作品ID4桁を抽出_(現在作品ID) ||
  作品ID4桁を抽出_(現在商品コード) ||
  作品ID4桁を抽出_(現在SKU) ||
  '';

let 行既存作品ID = '';
if (行既存作品ID候補) {
  const 候補ID = String(行既存作品ID候補).padStart(4, '0');
  const 候補WorksKey = 作品IDから既存WorksKey[候補ID] || '';

  if (候補WorksKey) {
    const 候補Data = 作品データ.keyToData[候補WorksKey] || {};

    const 現在比較キー = 作品比較キーを作る_(
      日本語タイトル,
      作者,
      原題,
      cfg,
      言語,
      カテゴリ
    );

    const 候補比較キー = 作品比較キーを作る_(
      候補Data.日本語タイトル || '',
      候補Data.作者 || '',
      候補Data.原題 || '',
      cfg,
      候補Data.言語 || '',
      候補Data.カテゴリ || ''
    );

    if (現在比較キー && 候補比較キー && 現在比較キー === 候補比較キー) {
      行既存作品ID = 候補ID;
    }
  }
}

      if (既存) {
        if (既存.作者) 作者 = 作者を名寄せ_(既存.作者, 作者別名マップ);
        if (!原題 && 既存.原題) 原題 = 正規化(既存.原題);
      }

      // その行が既に持っている作品IDを最優先
      if (行既存作品ID) {
        作品ID = String(行既存作品ID).padStart(4, '0');

        作品データ.keyToId[worksKey] = 作品ID;
        作品データ.keyToData[worksKey] = {
          日本語タイトル,
          作者,
          原題,
          言語,
          カテゴリ
        };

        const 同じ作品IDの既存Key = 作品IDから既存WorksKey[作品ID] || '';
        if (同じ作品IDの既存Key && 作品データ.keyToRow[同じ作品IDの既存Key]) {
          作品データ.keyToRow[worksKey] = 作品データ.keyToRow[同じ作品IDの既存Key];
          更新先WorksKey = 同じ作品IDの既存Key;
        } else if (作品データ.keyToRow[worksKey]) {
          更新先WorksKey = worksKey;
        } else {
          更新先WorksKey = worksKey;
          作品データ.keyToRow[worksKey] = null;
          作品データ.newRows.push([
            worksKey,
            作品ID,
            日本語タイトル,
            作者,
            原題を正規化(原題),
            '',
            '',
            '',
            '',
            ''
          ]);
        }

        作品IDから既存WorksKey[作品ID] = 更新先WorksKey;

      } else if (!作品ID) {
        const 候補 = 既存Works候補を探す_(
          日本語タイトル,
          作者,
          原題,
          言語,
          カテゴリ,
          作品データ,
          cfg
        );

        if (候補.found && 候補.作品ID) {
          作品ID = String(候補.作品ID).padStart(4, '0');

          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.keyToData[worksKey] = {
            日本語タイトル,
            作者,
            原題,
            言語,
            カテゴリ
          };

          if (候補.worksKey && 作品データ.keyToRow[候補.worksKey]) {
            作品データ.keyToRow[worksKey] = 作品データ.keyToRow[候補.worksKey];
            更新先WorksKey = 候補.worksKey;
          } else if (作品データ.keyToRow[worksKey]) {
            更新先WorksKey = worksKey;
          } else {
            更新先WorksKey = worksKey;
          }

          作品IDから既存WorksKey[作品ID] = 更新先WorksKey;

        } else {
          作品データ.maxId++;
          作品ID = String(作品データ.maxId).padStart(4, '0');
          // 発行した瞬間にハイウォーターを永続化（行削除後の同ID再発行を根絶）
          採番ハイウォーター更新(cfg, 作品データ.maxId);

          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.keyToData[worksKey] = {
            日本語タイトル,
            作者,
            原題,
            言語,
            カテゴリ
          };
          作品データ.keyToRow[worksKey] = null;

          作品データ.newRows.push([
            worksKey,
            作品ID,
            日本語タイトル,
            作者,
            原題を正規化(原題),
            '',
            '',
            '',
            '',
            ''
          ]);

          更新先WorksKey = worksKey;
          作品IDから既存WorksKey[作品ID] = worksKey;
        }

      } else {
        作品ID = String(作品ID).padStart(4, '0');

        const 同じ作品IDの既存Key = 作品IDから既存WorksKey[作品ID] || '';
        if (同じ作品IDの既存Key && 作品データ.keyToRow[同じ作品IDの既存Key]) {
          作品データ.keyToRow[worksKey] = 作品データ.keyToRow[同じ作品IDの既存Key];
          更新先WorksKey = 同じ作品IDの既存Key;
        } else if (作品データ.keyToRow[worksKey]) {
          更新先WorksKey = worksKey;
        } else {
          更新先WorksKey = worksKey;
        }
      }

      const 巻数 = 数値変換(取得生(cn.単巻数));
      const セット終了巻 = 数値変換(取得生(cn.セット終了));

      if (巻数 != null) {
        if (!作品データ.keyToVols[更新先WorksKey]) 作品データ.keyToVols[更新先WorksKey] = new Set();
        作品データ.keyToVols[更新先WorksKey].add(巻数);
      }

      if (セット終了巻 != null) {
        const セット開始巻 = 数値変換(取得生(cn.セット開始)) || 1;
        if (!作品データ.keyToVols[更新先WorksKey]) 作品データ.keyToVols[更新先WorksKey] = new Set();

        for (let v = セット開始巻; v <= セット終了巻; v++) {
          作品データ.keyToVols[更新先WorksKey].add(v);
        }
      }

      const SKU = SKUを生成(
        言語コード,
        形態コード,
        作品ID,
        カテゴリコード,
        取得生(cn.単巻数),
        取得生(cn.セット開始),
        取得生(cn.セット終了)
      );

      const タイトル = タイトルを段階生成(
        cfg,
        言語,
        カテゴリ,
        形態,
        日本語タイトル,
        取得生(cn.単巻数),
        取得生(cn.セット開始),
        取得生(cn.セット終了),
        作者,
        原題,
        取得(cn.特典メモ),
        形態マップ,
        クレジット種別
      );

      if (col作品ID) 変更map作品ID[sheetRow] = 作品ID;
      if (colSKU自動) 変更mapSKU自動[sheetRow] = SKU;
      if (col商品コード) 変更map商品コード[sheetRow] = SKU;
      if (colタイトル) 変更mapタイトル[sheetRow] = タイトル;
      if (colWorksタイトル) 変更mapWorksタイトル[sheetRow] = 商品タイトル;
      if (colステータス) 変更mapステータス[sheetRow] = '商品コード(発行済み確定)';
      if (col作者) 変更map作者[sheetRow] = 作者;
      if (col登録状況) 変更map登録状況[sheetRow] = '登録済み';

      発行数++;
    }

    列をまとめて書き込む(col作品ID, 変更map作品ID);
    列をまとめて書き込む(colSKU自動, 変更mapSKU自動);
    列をまとめて書き込む(col商品コード, 変更map商品コード);
    列をまとめて書き込む(colタイトル, 変更mapタイトル);
    列をまとめて書き込む(colWorksタイトル, 変更mapWorksタイトル);
    列をまとめて書き込む(colステータス, 変更mapステータス);
    列をまとめて書き込む(col作者, 変更map作者);
    列をまとめて書き込む(col登録状況, 変更map登録状況);

    作品データを更新_確定(作品シート, 作品データ, cfg);

    if (col発行チェック) {
      const 解除 = Array(最終行 - 1).fill([false]);
      sh.getRange(2, col発行チェック, 最終行 - 1, 1).setValues(解除);
    }

    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * ② 削除
 * ============================================================ */
function 削除を実行(cfg) {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(cfg.マスターシート名);
  if (!sh) return;

  const cn = cfg.列名;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ, cn);
  if (最終行 < 2) return;

  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 削除行 = [];

  for (let i = 0; i < 全データ.length; i++) {
    if (生値取得_(全データ[i], 列マップ, cn.発行チェック) === true) 削除行.push(i + 2);
  }

  if (削除行.length === 0) {
    ui.alert('チェックが入った行がありません');
    return;
  }

  if (ui.alert('確認', `${削除行.length}件を削除します。続行？`, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  行を一括削除(sh, 削除行);
  ui.alert(`✅ 削除完了: ${削除行.length}件`);
}

/* ============================================================
 * ③ 一括更新
 * ============================================================ */
function 一括更新を実行(cfg) {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '未登録行を再生成します（Works重複整理含む。作品IDの振り直しはしません＝永久ID）。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.マスターシート名);
  if (!sh) return;

  const cn = cfg.列名;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ, cn);
  if (最終行 < 2) {
    ui.alert('データがありません');
    return;
  }

  const masterSS = マスターSSを取得_(cfg);
  const 言語マップ = 言語マップをSSから取得_(masterSS, cfg);
  const カテゴリマップ = カテゴリマップをSSから取得_(masterSS, cfg);
  const 形態マップ = 形態マップをSSから取得_(masterSS, cfg);
  const 作者別名マップ = 作者別名マップを取得_(cfg);

  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターが見つかりません。マスターファイルIDを確認してください。');
    return;
  }

  const totalCols = sh.getLastColumn();
  const 全データ = sh.getRange(2, 1, 最終行 - 1, totalCols).getValues();
  const 作品シート = 作品シートを確保(ss, cfg);
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);

  const 商品名重複キーを作る_ = (原題商品タイトル) => {
    const k = キー用正規化_(原題商品タイトル);
    return k ? `ITEM||${k}` : '';
  };

  try {
    const 正規化結果 = WorksKey再正規化を実行(作品シート, cfg);
    const 孤立結果 = { 削除数: 0 };
    const ID結果 = WorksID振り直しを実行(作品シート, cfg);
    const 作品データ = 全作品データを読み込み(作品シート, cfg);

    const colWorksタイトル = 列番号安全取得_(列マップ, cn.Worksタイトル);
    const col作品ID = 列番号安全取得_(列マップ, cn.作品ID);
    const colSKU自動 = 列番号安全取得_(列マップ, cn.SKU自動);
    const col商品コード = 列番号安全取得_(列マップ, cn.商品コード);
    const colタイトル = 列番号安全取得_(列マップ, cn.タイトル);
    const colステータス = 列番号安全取得_(列マップ, cn.コードステータス);
    const col作者 = 列番号安全取得_(列マップ, cn.作者);
    const col発行チェック = 列番号安全取得_(列マップ, cn.発行チェック);

    const ISBNトラッカー = {};
    const 商品行キートラッカー = {};
    const 削除予定行 = [];

    const 変更map作品ID = {};
    const 変更mapSKU自動 = {};
    const 変更map商品コード = {};
    const 変更mapタイトル = {};
    const 変更mapWorksタイトル = {};
    const 変更mapステータス = {};
    const 変更map作者 = {};

    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const sheetRow = i + 2;

      const 取得 = (名前) => 値取得_(r, 列マップ, 名前);
      const 取得生 = (名前) => 生値取得_(r, 列マップ, 名前);
      const 取得列名直指定 = (列名文字列) => {
        const idx = 列マップ[列名文字列];
        return idx ? String(r[idx - 1] || '').trim() : '';
      };

      const 言語 = 取得(cn.言語);
      const カテゴリ = 取得(cn.カテゴリ);
      const 商品タイトル = 行の商品タイトルを取得_(r, 列マップ, cn, カテゴリ);
const 日本語タイトル = 取得(cn.日本語タイトル) || 商品タイトル;
let 作者 = 作者を名寄せ_(取得(cn.作者), 作者別名マップ);
let 原題 = 取得(cn.原題) || 商品タイトル;
      const 原題商品タイトル = cn.原題商品タイトル
        ? 取得(cn.原題商品タイトル)
        : 取得列名直指定('原題商品タイトル');
      const 形態 = 形態正規化_(取得(cn.形態));

      // 追加：アーティスト / 監督 / 出演 / 著 / 原作 / 関連
      const クレジット種別 = 取得(cn.クレジット種別);

      const ステータス = 取得生(cn.コードステータス);
      const 登録状況値 = String(取得生(cn.登録状況) || '').trim();
      const 確定行 = ステータス === '商品コード(発行済み確定)' || 登録状況値.startsWith('登録済');

      if (確定行) {
        if (日本語タイトル && 作者) {
          const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題, cfg);
          const 既存 = 作品データ.keyToData[worksKey];

          if (既存) {
            if (既存.作者) 作者 = 作者を名寄せ_(既存.作者, 作者別名マップ);
            if (!原題 && 既存.原題) 原題 = 正規化(既存.原題);
          }

          const 巻数 = 数値変換(取得生(cn.単巻数));
          const セット終了巻 = 数値変換(取得生(cn.セット終了));

          if (巻数 != null) {
            if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
            作品データ.keyToVols[worksKey].add(巻数);
          }

          if (セット終了巻 != null) {
            const セット開始巻 = 数値変換(取得生(cn.セット開始)) || 1;
            if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();

            for (let v = セット開始巻; v <= セット終了巻; v++) {
              作品データ.keyToVols[worksKey].add(v);
            }
          }
        }

        const ISBN確定 = 取得(cn.ISBN);
        const 博客來確定 = cn.博客來商品コード ? String(取得生(cn.博客來商品コード) || '').trim() : '';
        const 原題商品タイトル確定 = 原題商品タイトル;

        if (ISBN確定) {
          ISBNトラッカー['ISBN:' + ISBN確定] = true;
        } else if (博客來確定) {
          ISBNトラッカー['SITE:' + 博客來確定] = true;
        } else if (原題商品タイトル確定) {
          const 商品行キー = 商品名重複キーを作る_(原題商品タイトル確定);
          if (商品行キー) 商品行キートラッカー[商品行キー] = true;
        }

        if (col作者) 変更map作者[sheetRow] = 作者;
        continue;
      }

      if (!日本語タイトル || !言語 || !カテゴリ) continue;

      const ISBN = 取得(cn.ISBN);
      const 博客來コード = cn.博客來商品コード ? String(取得生(cn.博客來商品コード) || '').trim() : '';

      let 重複 = false;

      if (ISBN) {
        if (ISBNトラッカー['ISBN:' + ISBN]) 重複 = true;
        else ISBNトラッカー['ISBN:' + ISBN] = true;
      } else if (博客來コード) {
        if (ISBNトラッカー['SITE:' + 博客來コード]) 重複 = true;
        else ISBNトラッカー['SITE:' + 博客來コード] = true;
      } else if (原題商品タイトル) {
        const 商品行キー = 商品名重複キーを作る_(原題商品タイトル);
        if (商品行キー) {
          if (商品行キートラッカー[商品行キー]) 重複 = true;
          else 商品行キートラッカー[商品行キー] = true;
        }
      }

      if (重複) {
        削除予定行.push(sheetRow);
        if (colステータス) 変更mapステータス[sheetRow] = '【重複削除予定】';
        continue;
      }

      const 言語コード = 言語マップ[言語] || 'XX';
      const カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const 形態コード = 形態コードを取得_(形態, 形態マップ);

      let 作品ID = '';

      if (日本語タイトル && 作者) {
        const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題, cfg);
        const 既存 = 作品データ.keyToData[worksKey];

        if (既存) {
          if (既存.作者) 作者 = 作者を名寄せ_(既存.作者, 作者別名マップ);
          if (!原題 && 既存.原題) 原題 = 正規化(既存.原題);
        }

        作品ID = 作品データ.keyToId[worksKey];

        if (!作品ID) {
  const 候補 = 既存Works候補を探す_(
    日本語タイトル,
    作者,
    原題,
    言語,
    カテゴリ,
    作品データ,
    cfg
  );

  if (候補.found && 候補.作品ID) {
    作品ID = String(候補.作品ID).padStart(4, '0');

    作品データ.keyToId[worksKey] = 作品ID;
    作品データ.keyToData[worksKey] = {
      日本語タイトル,
      作者,
      原題,
      言語,
      カテゴリ
    };

  } else {
    作品データ.maxId++;
    作品ID = String(作品データ.maxId).padStart(4, '0');
    // 発行した瞬間にハイウォーターを永続化（行削除後の同ID再発行を根絶）
    採番ハイウォーター更新(cfg, 作品データ.maxId);

    作品データ.keyToId[worksKey] = 作品ID;
    作品データ.keyToData[worksKey] = {
      日本語タイトル,
      作者,
      原題,
      言語,
      カテゴリ
    };

    作品データ.newRows.push([
      worksKey,
      作品ID,
      日本語タイトル,
      作者,
      原題を正規化(原題),
      '',
      '',
      '',
      '',
      ''
    ]);
  }
} else {
  作品ID = String(作品ID).padStart(4, '0');
}
      }  // ← これを追加

      const SKU = SKUを段階生成(
        言語コード,
        形態コード,
        作品ID,
        カテゴリコード,
        取得生(cn.単巻数),
        取得生(cn.セット開始),
        取得生(cn.セット終了)
      );

      const タイトル = タイトルを段階生成(
        cfg,
        言語,
        カテゴリ,
        形態,
        日本語タイトル,
        取得生(cn.単巻数),
        取得生(cn.セット開始),
        取得生(cn.セット終了),
        作者,
        原題,
        取得(cn.特典メモ),
        形態マップ,
        クレジット種別
      );

      const newステータス = (日本語タイトル && 作者 && 言語 && カテゴリ)
        ? '商品コード(予約)'
        : '入力中...';

      変更map作品ID[sheetRow] = 作品ID;
      変更mapSKU自動[sheetRow] = SKU;
      変更map商品コード[sheetRow] = SKU;
      変更mapタイトル[sheetRow] = タイトル;
      変更mapWorksタイトル[sheetRow] = 商品タイトル;
      変更mapステータス[sheetRow] = newステータス;
      変更map作者[sheetRow] = 作者;
    }

        function 列をまとめて書き込む(colNum, valueMap) {
      if (!colNum) return;

      const rows = Object.keys(valueMap).map(Number).sort((a, b) => a - b);
      if (rows.length === 0) return;

      let start = rows[0];
      let buf = [[valueMap[rows[0]]]];

      for (let k = 1; k <= rows.length; k++) {
        if (k < rows.length && rows[k] === rows[k - 1] + 1) {
          buf.push([valueMap[rows[k]]]);
        } else {
          sh.getRange(start, colNum, buf.length, 1).setValues(buf);

          if (k < rows.length) {
            start = rows[k];
            buf = [[valueMap[rows[k]]]];
          }
        }
      }
    }

    列をまとめて書き込む(col作品ID, 変更map作品ID);
    列をまとめて書き込む(colSKU自動, 変更mapSKU自動);
    列をまとめて書き込む(col商品コード, 変更map商品コード);
    列をまとめて書き込む(colタイトル, 変更mapタイトル);
    列をまとめて書き込む(colWorksタイトル, 変更mapWorksタイトル);
    列をまとめて書き込む(colステータス, 変更mapステータス);
    列をまとめて書き込む(col作者, 変更map作者);

    作品データを更新_確定(作品シート, 作品データ, cfg);

    if (削除予定行.length > 0) {
      if (
        ui.alert(
          '重複削除',
          `重複行が${削除予定行.length}件あります。削除しますか？`,
          ui.ButtonSet.YES_NO
        ) === ui.Button.YES
      ) {
        行を一括削除(sh, 削除予定行);
      }
    }

    const 更新件数 = Object.keys(変更map作品ID).length;
    let msg = `✅ 一括更新完了: ${更新件数}件`;

    if (削除予定行.length > 0) msg += `\n重複削除: ${削除予定行.length}件`;
    if (正規化結果.統合数 > 0) msg += `\nWorks重複統合: ${正規化結果.統合数}件`;
    if (孤立結果.削除数 > 0) msg += `\nWorks孤立削除: ${孤立結果.削除数}件`;
    if (ID結果.変更数 > 0) msg += `\nID振り直し: ${ID結果.変更数}件`;

    if (col発行チェック) {
      const 解除 = Array(最終行 - 1).fill([false]);
      sh.getRange(2, col発行チェック, 最終行 - 1, 1).setValues(解除);
    }

    ui.alert(msg);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * ④ プルダウン更新
 * ============================================================ */
function プルダウン更新を実行(cfg) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.マスターシート名);

  if (!sh) {
    ui.alert('シートが見つかりません: ' + cfg.マスターシート名);
    return;
  }

  const cn = cfg.列名 || {};
  const 列マップ = 列番号を取得(sh);
  const masterSS = マスターSSを取得_(cfg);

  const 言語マップ = 言語マップをSSから取得_(masterSS, cfg);
  const カテゴリマップ = カテゴリマップをSSから取得_(masterSS, cfg);
  const 形態マップ = 形態マップをSSから取得_(masterSS, cfg);

  const 言語候補 = Object.keys(言語マップ || {});
  const カテゴリ候補 = Object.keys(カテゴリマップ || {});
  const 形態候補 = Object.keys((形態マップ && 形態マップ.コード) || {});

  const 最終行 = Math.max(sh.getMaxRows(), 1000);

  プルダウン設定_(sh, 列マップ, cn.登録状況, ['未登録', '登録済み'], 最終行);
  プルダウン設定_(sh, 列マップ, cn.言語, 言語候補, 最終行);
  プルダウン設定_(sh, 列マップ, cn.カテゴリ, カテゴリ候補, 最終行);
  プルダウン設定_(sh, 列マップ, cn.形態, 形態候補, 最終行);

  // 配送パターンはマスターがある場合だけ取得
  const 配送候補 = 配送パターン候補を取得_(masterSS);
  プルダウン設定_(sh, 列マップ, cn.配送パターン, 配送候補, 最終行);

  ui.alert('✅ プルダウン更新が完了しました');
}

function 配送パターン候補を取得_(masterSS) {
  const 候補シート名 = [
    '配送パターンマスター',
    '配送パターン',
    '配送マスター'
  ];

  for (const name of 候補シート名) {
    const sh = masterSS.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) continue;

    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .map(v => String(v || '').trim())
      .filter(Boolean);

    if (values.length > 0) return [...new Set(values)];
  }

  // マスターが無い場合の最低限フォールバック
  return [
    'ゆうパケ_(450)',
    '佐川'
  ];
}

/* ============================================================
 * マスター値取得
 * ============================================================ */
function 言語マップを取得(cfg) {
  return 言語マップをSSから取得_(マスターSSを取得_(cfg), cfg);
}
function カテゴリマップを取得(cfg) {
  return カテゴリマップをSSから取得_(マスターSSを取得_(cfg), cfg);
}
function 形態マップを取得(cfg) {
  return 形態マップをSSから取得_(マスターSSを取得_(cfg), cfg);
}

function 言語マップをSSから取得_(ss, cfg) {
  const map = {};
  const sh = ss.getSheetByName(cfg.言語マスター名 || '言語マスター');
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([名前, コード]) => {
    if (名前 && コード) map[正規化(名前)] = 正規化(コード);
  });
  return map;
}

function カテゴリマップをSSから取得_(ss, cfg) {
  const map = {};
  const sh = ss.getSheetByName(cfg.カテゴリマスター名 || 'カテゴリマスター');
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([名前, コード]) => {
    if (名前 && コード) map[正規化(名前)] = 正規化(コード);
  });
  return map;
}

function 形態マップをSSから取得_(masterSS, cfg) {
  const map = {
    コード: {},
    表示: {},
    特典テキスト: {},
    原題形態表記: {}
  };

  const sh = masterSS.getSheetByName(cfg.形態マスター名);
  if (!sh || sh.getLastRow() < 2) return map;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getLastColumn(), 5)).getValues();

  values.forEach(row => {
    const 形態名 = 正規化(row[0]);
    if (!形態名) return;

    map.コード[形態名] = 正規化(row[1]);
    map.特典テキスト[形態名] = 正規化(row[2]);
    map.表示[形態名] = 形態名;
    map.原題形態表記[形態名] = 正規化(row[4]);
  });

  return map;
}

function 言語マスターから値一覧_(ss, cfg) {
  const sh = ss.getSheetByName(cfg.言語マスター名 || '言語マスター');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(v => 正規化(v)).filter(Boolean);
}

function カテゴリマスターから値一覧_(ss, cfg) {
  const sh = ss.getSheetByName(cfg.カテゴリマスター名 || 'カテゴリマスター');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(v => 正規化(v)).filter(Boolean);
}

function 形態マスターから値一覧_(ss, cfg) {
  const sh = ss.getSheetByName(cfg.形態マスター名 || '形態マスター');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(v => 正規化(v)).filter(Boolean);
}

function 配送パターンマスターから値一覧_(ss) {
  const sh = ss.getSheetByName('配送パターンマスター');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(v => 正規化(v)).filter(Boolean);
}

/* ============================================================
 * データ最終行
 * ============================================================ */
function データがある最終行を取得(sh, 列マップ, cn) {
  const 候補列 = [
    cn.日本語タイトル,
    cn.作者,
    cn.原題,
    cn.商品コード,
    cn.作品ID
  ].map(name => 列番号安全取得_(列マップ, name)).filter(Boolean);

  if (候補列.length === 0) return sh.getLastRow();

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return lastRow;

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    for (const c of 候補列) {
      if (String(data[i][c - 1] || '').trim() !== '') return i + 2;
    }
  }
  return 1;
}

/* ============================================================
 * 行一括削除
 * ============================================================ */
function 行を一括削除(sh, 行番号配列) {
  if (!行番号配列 || 行番号配列.length === 0) return;
  const rows = [...行番号配列].sort((a, b) => b - a);
  for (const r of rows) {
    sh.deleteRow(r);
  }
}

/* ============================================================
 * 自己更新フラグ
 * ============================================================ */
function 自己更新中か_() {
  return PropertiesService.getScriptProperties().getProperty('KYOUTUU_UPDATING') === '1';
}
function 自己更新を開始_() {
  PropertiesService.getScriptProperties().setProperty('KYOUTUU_UPDATING', '1');
}
function 自己更新を終了_() {
  PropertiesService.getScriptProperties().deleteProperty('KYOUTUU_UPDATING');
}

/* ============================================================
 * SKU生成
 * ============================================================ */
function SKUを段階生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  if (!言語コード || !作品ID || !カテゴリコード) return '';

  const 巻数コード = 巻コードを作る_(単巻数, セット開始, セット終了);
  const 形態部 = 形態コード || '';

  return `${言語コード}${形態部}${String(作品ID).padStart(4, '0')}-${カテゴリコード}${巻数コード ? '-' + 巻数コード : ''}`;
}

function SKUを生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  return SKUを段階生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了);
}

function 巻コードを作る_(単巻数, セット開始, セット終了) {
  const 巻 = 数値変換(単巻数);
  if (巻 != null) return String(巻).padStart(2, '0');

  const 開 = 数値変換(セット開始);
  const 終 = 数値変換(セット終了);
  if (開 != null && 終 != null) {
    if (開 === 終) return String(開).padStart(2, '0');
    return String(開).padStart(2, '0') + String(終).padStart(2, '0');
  }
  return '';
}

function 巻表示を作る_(単巻数, セット開始, セット終了) {
  const 巻 = 数値変換(単巻数);
  if (巻 != null) return `${巻}巻`;

  const 開 = 数値変換(セット開始);
  const 終 = 数値変換(セット終了);

  if (開 != null && 終 != null) {
    if (開 === 終) return `${開}巻`;
    return `${開}〜${終}巻セット`;
  }

  return '';
}
/* ============================================================
 * タイトル生成
 * ============================================================ */
function タイトルを段階生成(
  cfg,
  言語,
  カテゴリ,
  形態,
  日本語タイトル,
  単巻数,
  セット開始,
  セット終了,
  作者,
  原題,
  特典メモ,
  形態マップ,
  クレジット種別 = ''
) {
  const parts = [];

  /**
   * 命名規則：
   * 言語 → カテゴリ → 形態 → 日本語タイトル → 巻数 → 作者 → 原題タイトル → 特典テキスト → 特典メモ
   */

  // 1. 言語
  if (言語) {
    parts.push(`${言語}版`);
  }

  // 2. カテゴリ
  if (カテゴリ) {
    parts.push(カテゴリ);
  }

  // 3. 形態
// プルダウンで選んだ形態名をそのままタイトルに入れる
if (形態 && 形態 !== '通常' && 形態 !== '通常版') {
  parts.push(形態);
}

  // 4. 日本語タイトル
  if (日本語タイトル) {
    parts.push(`『${日本語タイトル}』`);
  }

  // 5. 巻数
  const 巻表示 = 巻表示を作る_(単巻数, セット開始, セット終了);
  if (巻表示) {
    parts.push(巻表示);
  }

  // 6. 作者 / アーティスト / 監督 / 出演など
  const 作者表示する = !(cfg && cfg.タイトルに作者を表示 === false);

  let 作者ラベル = 正規化(クレジット種別);

  if (
    !作者ラベル &&
    cfg &&
    cfg.タイトル作者ラベルカテゴリ別 &&
    カテゴリ &&
    cfg.タイトル作者ラベルカテゴリ別[カテゴリ] != null
  ) {
    作者ラベル = String(cfg.タイトル作者ラベルカテゴリ別[カテゴリ]);
  }

  if (!作者ラベル && cfg && cfg.タイトル作者ラベル != null) {
    作者ラベル = String(cfg.タイトル作者ラベル);
  }

  // 本・まんが系のデフォルト
  if (!作者ラベル) {
    作者ラベル = '著';
  }

  if (作者表示する && 作者) {
    parts.push(`${作者ラベル}：${作者}`);
  }

  // 7. 原題タイトル
  const 原題表示する = !(cfg && cfg.タイトルに原題を表示 === false);

  if (原題表示する && 原題) {
    parts.push(`《${原題}》`);
  }

  // 8. 特典テキスト
  const 特典テキスト = 特典テキストを取得_(形態, 形態マップ);
if (特典テキスト) {
  parts.push(特典テキスト);
}

  // 9. 特典メモ
  if (特典メモ) {
    parts.push(特典メモ);
  }

  return parts
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 特典テキストを取得_(形態, 形態マップ) {
  const key = 正規化(形態);
  if (!key) return '';

  if (
    形態マップ &&
    形態マップ.特典テキスト &&
    形態マップ.特典テキスト[key] != null
  ) {
    return String(形態マップ.特典テキスト[key] || '').trim();
  }

  return '';
}

function 形態コードを取得_(形態, 形態マップ) {
  const key = 正規化(形態);
  if (!key) return '';

  if (
    形態マップ &&
    形態マップ.コード &&
    形態マップ.コード[key] != null
  ) {
    return String(形態マップ.コード[key] || '').trim();
  }

  return '';
}

function プルダウン設定_(sh, 列マップ, 列名, 候補リスト, 最終行, options) {
  options = options || {};

  if (!列名) return;

  const col = 列番号安全取得_(列マップ, 列名);
  if (!col) return;

  const values = [...new Set(
    (候補リスト || [])
      .map(v => String(v || '').trim())
      .filter(Boolean)
  )];

  if (values.length === 0) return;

  const 行数 = Math.max(最終行 - 1, 1);
  const range = sh.getRange(2, col, 行数, 1);

  /**
   * 重要：
   * 既にプルダウンがあるセルは上書きしない。
   * Googleスプレッドシートで手動設定した候補色が消えるのを防ぐ。
   */
  if (options.force !== true) {
    const validations = range.getDataValidations();

    let startOffset = null;
    let length = 0;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      .build();

    for (let i = 0; i < validations.length; i++) {
      const hasRule = validations[i][0] != null;

      if (!hasRule) {
        if (startOffset === null) {
          startOffset = i;
          length = 1;
        } else {
          length++;
        }
      }

      if ((hasRule || i === validations.length - 1) && startOffset !== null) {
        const endReachedWithMissing = !hasRule && i === validations.length - 1;
        const applyLength = endReachedWithMissing ? length : length;

        sh.getRange(2 + startOffset, col, applyLength, 1).setDataValidation(rule);

        startOffset = null;
        length = 0;
      }
    }

    return;
  }

  /**
   * force=true のときだけ全体を貼り直す。
   * ただし、この場合は候補色が消える可能性あり。
   */
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();

  range.setDataValidation(rule);
}

function 作品ID4桁を抽出_(v) {
  const s = 正規化(v).toUpperCase();
  if (!s) return '';

  if (/^\d{4}$/.test(s)) return s;
  if (/^\d+$/.test(s)) return String(parseInt(s, 10)).padStart(4, '0');

  // 例: TWRS0132-CM-02 / TWF0135-CM
  const m = s.match(/^[A-Z]{2}[A-Z]*?(\d{4})-/);
  return m ? m[1] : '';
}



