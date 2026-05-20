/* ============================================================
 * 台湾書籍系_共通.gs
 * 丸ごと置き換え用 完全フル版
 * ============================================================ */

/* ============================================================
 * 基本
 * ============================================================ */
function 台湾書籍系_正規化文字列_(v) {
  let s = String(v == null ? '' : v)
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/／/g, '/')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 全角英数字を半角に変換
  s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });

  return s;
}

function 台湾書籍系_列マップを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));

  const map = {};
  headers.forEach((h, i) => {
    if (h) map[h] = i + 1;
  });
  return map;
}

function 台湾書籍系_実列名を取得_(列, 候補配列) {
  for (const name of 候補配列) {
    const key = 台湾書籍系_正規化文字列_(name);
    if (key && 列[key]) return key;
  }
  return '';
}

function 台湾書籍系_数字2桁_(v) {
  const normalized = 台湾書籍系_正規化文字列_(v);
  const m = normalized.match(/\d+/);
  if (!m) return '';
  const n = parseInt(m[0], 10);
  if (isNaN(n)) return '';
  return String(n).padStart(2, '0');
}

/* ============================================================
 * Works 用
 * ============================================================ */
function 台湾書籍系_作品ヘッダー_() {
  return [
    'WorksKey',
    '作品ID',
    '日本語タイトル',
    '作者',
    '原題タイトル',
    '登録済み巻',
    '最新巻',
    '更新日時',
    '最新巻(予約込み)',
    '予約更新日時'
  ];
}

function 台湾書籍Worksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 'Works（書籍専用）';

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const headers = 台湾書籍系_作品ヘッダー_();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#cc0000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(1);
  const widths = [220, 100, 220, 180, 220, 120, 100, 150, 150, 150];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ Works（書籍専用）シートを作成しました');
}

/* ============================================================
 * 作品比較キー
 * 原題タイトルから巻数・版種を落として作品照合する
 * ============================================================ */
function 台湾書籍系_作品比較キー_(v) {
  return 台湾書籍系_正規化文字列_(v)
    .replace(/[！!？?]/g, '')
    .replace(/\s+/g, '')
    .replace(/(?:\d{1,3}|[０-９]{1,3})\s*(?:巻|册|集|部|話|期|號|号)?/gi, '')
    .replace(/(?:特装版|特裝版|首刷限定版|初版限定版|初回限定版|限定版|通常版)/gi, '')
    .trim();
}

/* ============================================================
 * 商品比較キー
 * 商品重複判定用
 * ============================================================ */
function 台湾書籍系_商品比較キー_(v) {
  return 台湾書籍系_正規化文字列_(v)
    .replace(/[！!？?]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

/* ============================================================
 * 形態
 * ============================================================ */
function 台湾書籍系_形態を正規化_(v) {
  return 台湾書籍系_正規化文字列_(v);
}

function 台湾書籍系_形態コードを生成_(形態, 設定) {
  const v = 台湾書籍系_形態を正規化_(形態);
  if (!v) return '';

  const map = 台湾書籍系_形態マスターを取得_(設定);
  if (map && map.nameToCode && map.nameToCode[v] !== undefined) {
    return map.nameToCode[v];
  }

  // 念のための後方互換
  if (v === '初版限定版') return 'F';
  if (v === '特装版') return 'S';

  return '';
}
/* ============================================================
 * 表示生成
 * ============================================================ */
function 台湾書籍系_マスターSSを取得_(設定) {
  try {
    if (設定 && 設定.マスターファイルID) {
      return SpreadsheetApp.openById(設定.マスターファイルID);
    }
  } catch (err) {
    console.log('台湾書籍系_マスターSSを取得_ error: ' + err);
  }
  return SpreadsheetApp.getActive();
}

function 台湾書籍系_言語マスターを取得_(設定) {
  const ss = 台湾書籍系_マスターSSを取得_(設定);
  const sh = ss.getSheetByName('言語マスター');
  const result = {
    nameToCode: {},
    codeToName: {}
  };
  if (!sh || sh.getLastRow() < 2) return result;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  values.forEach(([name, code]) => {
    const n = 台湾書籍系_正規化文字列_(name);
    const c = 台湾書籍系_正規化文字列_(code).toUpperCase();
    if (!n) return;
    if (c) {
      result.nameToCode[n] = c;
      result.codeToName[c] = n;
    } else {
      result.nameToCode[n] = '';
    }
  });
  return result;
}

function 台湾書籍系_カテゴリマスターを取得_(設定) {
  const ss = 台湾書籍系_マスターSSを取得_(設定);
  const sh = ss.getSheetByName('カテゴリマスター');
  const result = {
    nameToCode: {},
    codeToName: {}
  };
  if (!sh || sh.getLastRow() < 2) return result;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  values.forEach(([name, code]) => {
    const n = 台湾書籍系_正規化文字列_(name);
    const c = 台湾書籍系_正規化文字列_(code).toUpperCase();
    if (!n) return;
    if (c) {
      result.nameToCode[n] = c;
      result.codeToName[c] = n;
    } else {
      result.nameToCode[n] = '';
    }
  });
  return result;
}

function 台湾書籍系_形態マスターを取得_(設定) {
  const ss = 台湾書籍系_マスターSSを取得_(設定);
  const sh = ss.getSheetByName('形態マスターシート');
  const result = {
    nameToCode: {},
    codeToName: {},
    nameToBonusText: {}
  };
  if (!sh || sh.getLastRow() < 2) return result;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  values.forEach(([name, code, bonusText]) => {
    const n = 台湾書籍系_正規化文字列_(name);
    const c = 台湾書籍系_正規化文字列_(code).toUpperCase();
    const b = 台湾書籍系_正規化文字列_(bonusText);
    if (!n) return;

    result.nameToCode[n] = c || '';
    if (c) result.codeToName[c] = n;
    if (b) result.nameToBonusText[n] = b;
  });
  return result;
}

function 台湾書籍系_言語表示を生成_(言語, 設定) {
  const v = 台湾書籍系_正規化文字列_(言語);
  if (!v) return '';

  const map = 台湾書籍系_言語マスターを取得_(設定);
  const upper = v.toUpperCase();

  if (map.codeToName[upper]) return `${map.codeToName[upper]}版`;
  if (map.nameToCode[v] !== undefined) return `${v}版`;

  return `${v}版`;
}

function 台湾書籍系_カテゴリ表示を生成_(シート名, カテゴリ, 設定) {
  const v = 台湾書籍系_正規化文字列_(カテゴリ);
  const map = 台湾書籍系_カテゴリマスターを取得_(設定);

  if (v) {
    const upper = v.toUpperCase();
    if (map.codeToName[upper]) return map.codeToName[upper];
    if (map.nameToCode[v] !== undefined) return v;
    return v;
  }

  return 台湾書籍系_正規化文字列_(シート名) === '台湾まんが' ? 'まんが' : '書籍';
}

function 台湾書籍系_巻表示を生成_(単巻数, セット開始, セット終了) {
  const 単 = 台湾書籍系_数字2桁_(単巻数);
  const 開 = 台湾書籍系_数字2桁_(セット開始);
  const 終 = 台湾書籍系_数字2桁_(セット終了);

  if (開 && 終) {
    if (開 === 終) return `第${開}巻`;
    return `${開}-${終}巻`;
  }
  if (単) return `第${単}巻`;
  return '';
}

function 台湾書籍系_タイトルを生成_(シート名, 値, 設定) {
  const norm = (v) => 台湾書籍系_正規化文字列_(v || '');

  const stripJpTitle = (title) => {
    let s = norm(title);
    if (!s) return '';
    return s
      .replace(/\s*[（(]\s*\d{1,3}\s*[）)]\s*$/u, '')
      .replace(/\s+第?\d{1,3}\s*巻\s*$/u, '')
      .replace(/\s+\d{1,3}\s*巻\s*$/u, '')
      .trim();
  };

  const shapeMap = 台湾書籍系_形態マスターを取得_(設定);

  const 言語表記 = 台湾書籍系_言語表示を生成_(値.言語, 設定);
  const カテゴリ表記 = 台湾書籍系_カテゴリ表示を生成_(シート名, 値.カテゴリ, 設定);
  const 形態表記 = 台湾書籍系_形態を正規化_(値.形態);
  const 日本語タイトル = stripJpTitle(値.日本語タイトル || 値.原題タイトル || '');
  let 巻表示 = 台湾書籍系_巻表示を生成_(値.単巻数, 値.セット開始, 値.セット終了);
if (!巻表示) {
  // 単巻数が空の場合は原題などから自動抽出
  const 抽出巻数 = 台湾書籍系_巻数を抽出_(値);
  if (抽出巻数) 巻表示 = `第${抽出巻数}巻`;
}
  const 作者 = norm(値.作者);

  // 原題商品タイトル優先、なければ原題タイトル
  const 元タイトル = norm(値.原題商品タイトル || 値.原題タイトル);

  // 特典文は「特典メモ」と「形態マスターの特典テキスト」を両方出す
  const 特典メモ = norm(値.特典メモ);
  const 形態特典文 = (形態表記 && shapeMap.nameToBonusText[形態表記])
    ? norm(shapeMap.nameToBonusText[形態表記])
    : '';

  if (!日本語タイトル) return '';

  const parts = [];

  if (言語表記) parts.push(言語表記);

  // 形態はカテゴリの直後にカッコ付きで表示
  // 例：台湾版 まんが(予約限定特装版)
  if (カテゴリ表記) {
    if (形態表記 &&形態表記 !== '通常版') {
      parts.push(`${カテゴリ表記}(${形態表記})`);
    } else {
      parts.push(カテゴリ表記);
    }
  } else if (形態表記 && 形態表記 !== '通常版') {
    parts.push(`(${形態表記})`);
  }

  parts.push(`『${日本語タイトル}』`);

  if (巻表示) parts.push(巻表示);
  if (作者) parts.push(`著：${作者}`);
  if (元タイトル) parts.push(元タイトル);

  // 両方入れる。ただし同文なら1回だけ
  if (形態特典文) parts.push(形態特典文);
  if (特典メモ && 特典メモ !== 形態特典文) parts.push(特典メモ);

  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/* ============================================================
 * Works 検索
 * 原題タイトルベース
 * 原題商品タイトルは使わない
 * ============================================================ */
function 台湾書籍系_Worksから取得_(設定, 原題タイトル, 日本語タイトル = '', 作者 = '') {
  const key = 台湾書籍系_正規化文字列_(原題タイトル);
  // ★追加
  
  const 作品キー = 台湾書籍系_作品比較キー_(原題タイトル);
  const jpKey = 台湾書籍系_正規化文字列_(日本語タイトル);
  const authorKey = 台湾書籍系_正規化文字列_(作者);

  if (!key && !作品キー && !jpKey) return null;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return null;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));

  const col = {};
  headers.forEach((h, i) => {
    if (h) col[h] = i;
  });

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
// 1) 完全一致の直前に
  
  // 1) 原題タイトル 完全一致
  for (const row of data) {
    const raw = 台湾書籍系_正規化文字列_(row[col['原題タイトル']]);
    if (raw && raw === key) {
      return {
        作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者: col['作者'] != null ? row[col['作者']] : ''
      };
    }
  }

  // 2) 原題タイトル 作品比較キー一致
  for (const row of data) {
    const raw = 台湾書籍系_正規化文字列_(row[col['原題タイトル']]);
    const raw作品キー = 台湾書籍系_作品比較キー_(raw);
    if (raw作品キー && 作品キー && raw作品キー === 作品キー) {
      return {
        作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者: col['作者'] != null ? row[col['作者']] : ''
      };
    }
  }

  // 3) 日本語タイトル + 作者 fallback
  if (jpKey && authorKey && col['日本語タイトル'] != null && col['作者'] != null) {
    for (const row of data) {
      const jp = 台湾書籍系_正規化文字列_(row[col['日本語タイトル']]);
      const au = 台湾書籍系_正規化文字列_(row[col['作者']]);
      if (jp === jpKey && au === authorKey) {
        return {
          作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
          日本語タイトル: jp,
          作者: au
        };
      }
    }
  }

  return null;
}

/* ============================================================
 * 商品重複チェックキー
 * ISBN → サイト商品コード → 原題商品タイトル
 * ============================================================ */
function 台湾書籍系_重複チェックキーを生成_(値) {
  const isbn = 台湾書籍系_正規化文字列_(値.ISBN).replace(/[^0-9Xx]/g, '');
  if (isbn) return `ISBN||${isbn}`;

  const siteCode = 台湾書籍系_正規化文字列_(値.サイト商品コード);
  if (siteCode) return `SITE||${siteCode}`;

  const rawItem = 台湾書籍系_商品比較キー_(値.原題商品タイトル);
  if (rawItem) return `ITEM||${rawItem}`;

  return '';
}

/* ============================================================
 * 粗利益率
 * ============================================================ */
function 台湾書籍系_粗利益率を計算_(売価, 原価) {
  const 売 = parseFloat(売価 || 0);
  const 原 = parseFloat(原価 || 0);
  if (!(売 > 0 && 原 > 0)) return null;

  let レート = 0;
  try {
    if (typeof _kyoutuu !== 'undefined' && _kyoutuu.為替レートを取得_) {
      レート = _kyoutuu.為替レートを取得_('TWD');
    }
  } catch (_) {
    レート = 0;
  }
  if (!(レート > 0)) return null;

  const 送料率 = 1.3;       // 原価×30%
  const 販売コスト率 = 0.22; // Yahoo9% + カード3% + 消費税10%

  return Math.round(((売 - 原 * レート * 送料率 - 売 * 販売コスト率) / 売) * 1000) / 1000;
}

/* ============================================================
 * 採番ヘルパー
 * ============================================================ */
function 台湾書籍系_関数を探す_(候補名配列) {
  for (const name of 候補名配列) {
    try {
      const fn = globalThis[name];
      if (typeof fn === 'function') return fn;
    } catch (_) {}
  }
  return null;
}

function 台湾書籍系_言語コードを生成_(言語) {
  const v = 台湾書籍系_正規化文字列_(言語).toUpperCase();
  if (!v) return 'TW';
  if (v === '台湾' || v === 'TW') return 'TW';
  if (v === '中国' || v === 'CN') return 'CN';
  if (v === '韓国' || v === 'KR') return 'KR';
  if (v === '香港' || v === 'HK') return 'HK';
  if (v === '日本' || v === 'JP') return 'JP';
  return v.replace(/[^A-Z0-9]/g, '').slice(0, 2) || 'TW';
}

function 台湾書籍系_カテゴリコードを生成_(カテゴリ, シート名) {
  const v = 台湾書籍系_正規化文字列_(カテゴリ);

  if (/まんが|漫画|コミック/i.test(v)) return 'CM';
  if (/小説|ノベル/i.test(v)) return 'NV';
  if (/アート|画集|設定集|資料集/i.test(v)) return 'ART';
  if (/雑誌/i.test(v)) return 'MZ';
  if (/グッズ/i.test(v)) return 'GD';

  if (台湾書籍系_正規化文字列_(シート名) === '台湾まんが') return 'CM';

  const raw = v.toUpperCase().replace(/[^A-Z0-9]/g, '');
  return raw ? raw.slice(0, 3) : 'BK';
}

function 台湾書籍系_作品ID4桁を取得_(作品ID) {
  const s = 台湾書籍系_正規化文字列_(作品ID);
  if (!s) return '';

  if (/^\d{4}$/.test(s)) return s;
  if (/^\d+$/.test(s)) return String(parseInt(s, 10)).padStart(4, '0');

  return '';
}

function 台湾書籍系_巻数を抽出_(値) {
  const 単巻 = 台湾書籍系_数字2桁_(値.単巻数);
  if (単巻) return 単巻;

  const 開始 = 台湾書籍系_数字2桁_(値.セット開始);
  const 終了 = 台湾書籍系_数字2桁_(値.セット終了);
  if (開始 && 終了) {
    if (開始 === 終了) return 開始;
    return `${開始}${終了}`;
  }

  const texts = [
    値.原題商品タイトル,
    値.原題タイトル,
    値.日本語タイトル
  ].map(v => 台湾書籍系_正規化文字列_(v)).filter(Boolean);

  for (const text of texts) {
    const patterns = [
      /(?:^|[^\d])(\d{1,3})(?=\s*(?:巻|册|集|部|話|期|號|号))/i,
      /(?:^|[^\d])(\d{1,3})(?=\s*(?:特装版|特裝版|首刷限定版|初版限定版|初回限定版|限定版|通常版|特装|特裝|首刷|初版|版|特典|$))/i
    ];

    for (const re of patterns) {
      const m = text.match(re);
      if (m) return String(parseInt(m[1], 10)).padStart(2, '0');
    }

    const nums = text.match(/\d{1,3}/g);
    if (nums && nums.length) {
      const last = parseInt(nums[nums.length - 1], 10);
      if (!isNaN(last)) return String(last).padStart(2, '0');
    }
  }

  return '';
}

function 台湾書籍系_コードから作品ID4桁を取得_(v) {
  const s = 台湾書籍系_正規化文字列_(v).toUpperCase();
  if (!s) return '';

  // 例:
  // TW0002-CM-16
  // TWF0002-CM-17
  // TWS0007-CM-0304
  const m = s.match(/^[A-Z]{2}[A-Z]*?(\d{4})-/);
  return m ? m[1] : '';
}

function 台湾書籍系_行から既存作品IDを取得_(値) {
  return (
    台湾書籍系_作品ID4桁を取得_(値.作品ID) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.商品コード) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.SKU自動) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.親コード) ||
    ''
  );
}

function 台湾書籍系_親コードベースを組み立て_(作品ID, 値, シート名, 設定) {
  const langCode = 台湾書籍系_言語コードを生成_(値.言語);
  const shapeCode = 台湾書籍系_形態コードを生成_(値.形態, 設定);
  const workId4 = 台湾書籍系_作品ID4桁を取得_(作品ID);
  const catCode = 台湾書籍系_カテゴリコードを生成_(値.カテゴリ, シート名);

  if (!workId4) return '';

  return `${langCode}${shapeCode}${workId4}-${catCode}`;
}

function 台湾書籍系_最終SKUを組み立て_(baseCode, 値) {
  const volumeCode = 台湾書籍系_巻数を抽出_(値);
  if (!baseCode) return '';
  if (!volumeCode) return baseCode;
  return `${baseCode}-${volumeCode}`;
}

function 台湾書籍系_親コードを生成_共通_(ctx) {
  const { sheet, 値, 実列名, rowValues, 列, 設定 } = ctx;

  const 作品ID列名 = 実列名 && 実列名.作品ID ? 実列名.作品ID : '';
  const raw作品ID = 作品ID列名 ? rowValues[列[作品ID列名] - 1] : 値.作品ID;
  const 作品ID = 台湾書籍系_作品ID4桁を取得_(raw作品ID);

  return 台湾書籍系_親コードベースを組み立て_(
    作品ID,
    値,
    sheet.getName(),
    設定
  );
}

function 台湾書籍系_SKUを生成_共通_(ctx) {
  const { sheet, 値, 親コード, 実列名, rowValues, 列, 設定 } = ctx;

  let base = 台湾書籍系_正規化文字列_(親コード);

  if (!base) {
    const 作品ID列名 = 実列名 && 実列名.作品ID ? 実列名.作品ID : '';
    const raw作品ID = 作品ID列名 ? rowValues[列[作品ID列名] - 1] : 値.作品ID;
    const 作品ID = 台湾書籍系_作品ID4桁を取得_(raw作品ID);

    base = 台湾書籍系_親コードベースを組み立て_(
      作品ID,
      値,
      sheet.getName(),
      設定
    );
  }

  return 台湾書籍系_最終SKUを組み立て_(base, 値);
}

/* ============================================================
 * 不足項目
 * ============================================================ */
function 台湾書籍系_不足項目を返す_(シート名, 値, 設定) {
  const lacks = [];
  if (!台湾書籍系_正規化文字列_(値.言語)) lacks.push('言語');
  if (!台湾書籍系_正規化文字列_(値.日本語タイトル || 値.原題タイトル)) lacks.push('タイトル');

  // ★修正：形態は必須から外す（空＝通常版として扱う）
  // if (!台湾書籍系_形態を正規化_(値.形態)) lacks.push('形態');

  const カテゴリ = 台湾書籍系_カテゴリ表示を生成_(シート名, 値.カテゴリ, 設定);
  if (!カテゴリ) lacks.push('カテゴリ');
  if (!台湾書籍系_作品ID4桁を取得_(値.作品ID)) lacks.push('作品ID');
  return lacks;
}

/* ============================================================
 * 親コード / SKU 試行生成
 * ============================================================ */
function 台湾書籍系_親コードSKUを試行生成_(sh, row, 設定, 列, rowValues, 値, 実列名) {
  let changed = false;

  デバッグログ出力_('書籍系_親コードSKU試行開始', {
    sheet: sh.getName(),
    row,
    商品コード列名: 実列名.商品コード,
    SKU列名: 実列名.SKU自動
  });

  const 親コード関数 = 台湾書籍系_関数を探す_([
    `${sh.getName()}_親コードを生成_`,
    '台湾書籍系_親コードを生成_共通_'
  ]);

  const SKU関数 = 台湾書籍系_関数を探す_([
    `${sh.getName()}_SKUを生成_`,
    '台湾書籍系_SKUを生成_共通_'
  ]);

  デバッグログ出力_('書籍系_生成関数確認', {
    sheet: sh.getName(),
    row,
    親コード関数あり: !!親コード関数,
    SKU関数あり: !!SKU関数
  });

  const 商品コード列名 = 実列名.商品コード;
  const SKU列名 = 実列名.SKU自動;
  const ステータス列名 = 実列名.コードステータス;

  const 現在親コード = 商品コード列名
    ? 台湾書籍系_正規化文字列_(rowValues[列[商品コード列名] - 1])
    : '';
  const 現在SKU = SKU列名
    ? 台湾書籍系_正規化文字列_(rowValues[列[SKU列名] - 1])
    : '';

  // 生成関数が両方ないときだけ終了
  if (!親コード関数 && !SKU関数) {
    if (ステータス列名 && !現在親コード && !現在SKU) {
      rowValues[列[ステータス列名] - 1] = '生成関数未定義';
      changed = true;
    }

    デバッグログ出力_('書籍系_生成結果', {
      sheet: sh.getName(),
      row,
      親コード: '',
      SKU: '',
      ステータス: ステータス列名 ? rowValues[列[ステータス列名] - 1] : ''
    });

    return changed;
  }

  // 親コードは毎回再生成
  if (親コード関数 && 商品コード列名) {
    try {
      const code = 親コード関数({
        sheet: sh,
        row,
        設定,
        列,
        値,
        実列名,
        rowValues
      });

      if (code) {
        rowValues[列[商品コード列名] - 1] = code;
        changed = true;
      }
    } catch (err) {
      if (ステータス列名) {
        rowValues[列[ステータス列名] - 1] = `親コード生成エラー: ${err}`;
      }
      changed = true;

      デバッグログ出力_('書籍系_親コード生成エラー', {
        sheet: sh.getName(),
        row,
        message: String(err),
        stack: err && err.stack ? String(err.stack) : ''
      });
    }
  }

  const 生成後親コード = 商品コード列名
    ? 台湾書籍系_正規化文字列_(rowValues[列[商品コード列名] - 1])
    : '';

  // SKUも毎回再生成
  if (SKU関数 && SKU列名) {
    try {
      const sku = SKU関数({
        sheet: sh,
        row,
        設定,
        列,
        値,
        親コード: 生成後親コード,
        実列名,
        rowValues
      });

      if (sku) {
        rowValues[列[SKU列名] - 1] = sku;
        changed = true;
      }
    } catch (err) {
      if (ステータス列名) {
        rowValues[列[ステータス列名] - 1] = `SKU生成エラー: ${err}`;
      }
      changed = true;

      デバッグログ出力_('書籍系_SKU生成エラー', {
        sheet: sh.getName(),
        row,
        message: String(err),
        stack: err && err.stack ? String(err.stack) : ''
      });
    }
  }

  // 最終SKUを商品コード列にも上書き
  const 生成後SKU = SKU列名
    ? 台湾書籍系_正規化文字列_(rowValues[列[SKU列名] - 1])
    : '';

  if (商品コード列名 && 生成後SKU) {
    rowValues[列[商品コード列名] - 1] = 生成後SKU;
    changed = true;
  }

  デバッグログ出力_('書籍系_生成結果', {
    sheet: sh.getName(),
    row,
    親コード: 生成後親コード,
    SKU: SKU列名 ? rowValues[列[SKU列名] - 1] : '',
    商品コード: 商品コード列名 ? rowValues[列[商品コード列名] - 1] : '',
    ステータス: ステータス列名 ? rowValues[列[ステータス列名] - 1] : ''
  });

  return changed;
}

/* ============================================================
 * 1行補完
 * ============================================================ */
function 台湾書籍系_商品コードだけから作品IDを取得_(値) {
  return (
    台湾書籍系_コードから作品ID4桁を取得_(値.商品コード) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.SKU自動) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.親コード) ||
    ''
  );
}

function 台湾書籍系_1行補完_共通_(sh, row, 設定, options = {}) {
  if (!sh || row < 2) return;

  if (options.skipFlush !== true) {
    SpreadsheetApp.flush();
  }

  const cn = 設定.列名 || {};
  const 列 = options.列マップ || 台湾書籍系_列マップを取得_(sh);
  const lastCol = options.lastCol || sh.getLastColumn();
  const rowValues = options.rowValues || sh.getRange(row, 1, 1, lastCol).getValues()[0];

  const 実列名 = {
    商品コード: 台湾書籍系_実列名を取得_(列, [cn.商品コード, '親コード', '商品コード', '商品コード(SKU)', '商品コード（SKU）']),
    タイトル: 台湾書籍系_実列名を取得_(列, [cn.タイトル, 'タイトル']),
    作者: 台湾書籍系_実列名を取得_(列, [cn.作者, '作者']),
    日本語タイトル: 台湾書籍系_実列名を取得_(列, [cn.日本語タイトル, '日本語タイトル']),
    原題: 台湾書籍系_実列名を取得_(列, [cn.原題, '原題タイトル']),
    原題商品タイトル: 台湾書籍系_実列名を取得_(列, [cn.原題商品タイトル, '原題商品タイトル']),
    形態: 台湾書籍系_実列名を取得_(列, [cn.形態, '形態', '形態(通常/初回限定/特装)', '形態（通常/初回限定/特装）']),
    言語: 台湾書籍系_実列名を取得_(列, [cn.言語, '言語']),
    カテゴリ: 台湾書籍系_実列名を取得_(列, [cn.カテゴリ, 'カテゴリ']),
    単巻数: 台湾書籍系_実列名を取得_(列, [cn.単巻数, '単巻数']),
    セット開始: 台湾書籍系_実列名を取得_(列, [cn.セット開始, 'セット巻数開始番号']),
    セット終了: 台湾書籍系_実列名を取得_(列, [cn.セット終了, 'セット巻数終了番号']),
    特典メモ: 台湾書籍系_実列名を取得_(列, [cn.特典メモ, '特典メモ']),
    ISBN: 台湾書籍系_実列名を取得_(列, [cn.ISBN, 'ISBN']),
    作品ID: 台湾書籍系_実列名を取得_(列, [cn.作品ID, '作品ID(W)(自動)', '作品ID(W)（自動）', '作品(W)（自動）']),
    SKU自動: 台湾書籍系_実列名を取得_(列, [cn.SKU自動, 'SKU(自動)', 'SKU（自動）', '商品コード(SKU)', '商品コード（SKU）']),
    コードステータス: 台湾書籍系_実列名を取得_(列, [cn.コードステータス, '商品コードステータス']),
    発行チェック: 台湾書籍系_実列名を取得_(列, [cn.発行チェック, '発番発行']),
    登録状況: 台湾書籍系_実列名を取得_(列, [cn.登録状況, '登録状況']),
    サイト商品コード: 台湾書籍系_実列名を取得_(列, [cn.サイト商品コード, cn.博客來商品コード, 'サイト商品コード', '博客來商品コード']),
    売価: 台湾書籍系_実列名を取得_(列, [cn.売価, '売価']),
    原価: 台湾書籍系_実列名を取得_(列, [cn.原価, '原価']),
    粗利益率: 台湾書籍系_実列名を取得_(列, [cn.粗利益率, '粗利益率']),
    重複チェックキー: 台湾書籍系_実列名を取得_(列, [cn.重複チェックキー, '重複チェックキー'])
  };

  const g = (列名) => {
    if (!列名 || !列[列名]) return '';
    return rowValues[列[列名] - 1];
  };

  const s = (列名, 値) => {
    if (!列名 || !列[列名]) return false;
    const idx = 列[列名] - 1;
    const before = rowValues[idx];
    if (String(before) === String(値)) return false;
    rowValues[idx] = 値;
    return true;
  };

  let changed = false;

  let 値 = {
    作者: g(実列名.作者),
    日本語タイトル: g(実列名.日本語タイトル),
    原題タイトル: g(実列名.原題),
    原題商品タイトル: g(実列名.原題商品タイトル),
    形態: 台湾書籍系_形態を正規化_(g(実列名.形態)),
    言語: g(実列名.言語),
    カテゴリ: g(実列名.カテゴリ),
    単巻数: g(実列名.単巻数),
    セット開始: g(実列名.セット開始),
    セット終了: g(実列名.セット終了),
    特典メモ: g(実列名.特典メモ),
    ISBN: g(実列名.ISBN),
    サイト商品コード: g(実列名.サイト商品コード),
    作品ID: g(実列名.作品ID),
    登録状況: g(実列名.登録状況),
    商品コード: g(実列名.商品コード),
    親コード: g(実列名.商品コード),
    SKU自動: g(実列名.SKU自動)
  };

  const 登録済み行 = /^登録済/.test(台湾書籍系_正規化文字列_(値.登録状況));
  const 現在作品ID = 台湾書籍系_作品ID4桁を取得_(値.作品ID);
  const 商品コード由来ID = 台湾書籍系_商品コードだけから作品IDを取得_(値);

  /*
   * 登録済み行の鉄則：
   * 商品コード / SKU に含まれる作品IDを正とする。
   * Works側のIDで上書きしない。
   * 親コード / SKU / 商品コードステータスも再生成しない。
   */
  if (登録済み行) {
    if (商品コード由来ID && 現在作品ID !== 商品コード由来ID) {
      値.作品ID = 商品コード由来ID;

      if (実列名.作品ID) {
        changed = s(実列名.作品ID, 商品コード由来ID) || changed;
      }

      デバッグログ出力_('登録済み行_商品コード由来IDを優先', {
        row,
        現在作品ID,
        商品コード由来ID,
        商品コード: 値.商品コード,
        SKU自動: 値.SKU自動,
        親コード: 値.親コード
      });
    }

    if (changed) {
      const 粗利益率列 = 実列名.粗利益率 ? 列[実列名.粗利益率] : 0;

      台湾書籍系_1行を書き戻す_粗利益率除外_(
        sh,
        row,
        lastCol,
        rowValues,
        粗利益率列
      );

      if (実列名.作品ID) {
        sh.getRange(row, 列[実列名.作品ID]).setNumberFormat('@');
      }
    }

    return;
  }

  /*
   * ここから下は未登録行だけ。
   * 未登録行は従来どおり Works 照合・タイトル生成・SKU生成を行う。
   */
  if (!現在作品ID) {
    const 既存ID = 台湾書籍系_行から既存作品IDを取得_(値);

    if (既存ID) {
      値.作品ID = 既存ID;

      if (実列名.作品ID) {
        changed = s(実列名.作品ID, 既存ID) || changed;
      }

      デバッグログ出力_('作品IDを既存情報から復元', {
        row,
        既存ID,
        商品コード: 値.商品コード,
        SKU自動: 値.SKU自動,
        親コード: 値.親コード
      });
    }
  }

  const Works新規作成を許可 = options.Works新規作成 === true;

  const works = Works新規作成を許可
    ? 台湾書籍系_Worksを取得または作成_(設定, 値)
    : 台湾書籍系_Worksから取得_(設定, 値.原題タイトル, 値.日本語タイトル, 値.作者);

  if (works) {
    const worksID = 台湾書籍系_作品ID4桁を取得_(works.作品ID);
    const 現在ID = 台湾書籍系_作品ID4桁を取得_(値.作品ID);

    // 未登録行だけ、作品IDが空ならWorks IDを入れる
    if (worksID && !現在ID) {
      changed = s(実列名.作品ID, worksID) || changed;
      値.作品ID = worksID;
    }

    if (!台湾書籍系_正規化文字列_(値.日本語タイトル) && 台湾書籍系_正規化文字列_(works.日本語タイトル)) {
      changed = s(実列名.日本語タイトル, works.日本語タイトル) || changed;
      値.日本語タイトル = works.日本語タイトル;
    }

    if (!台湾書籍系_正規化文字列_(値.作者) && 台湾書籍系_正規化文字列_(works.作者)) {
      changed = s(実列名.作者, works.作者) || changed;
      値.作者 = works.作者;
    }

    if (!台湾書籍系_正規化文字列_(値.原題タイトル) && 台湾書籍系_正規化文字列_(works.原題タイトル)) {
      changed = s(実列名.原題, works.原題タイトル) || changed;
      値.原題タイトル = works.原題タイトル;
    }
  }

  const 正規化作品ID = 台湾書籍系_作品ID4桁を取得_(値.作品ID);
  if (正規化作品ID && String(値.作品ID) !== 正規化作品ID) {
    changed = s(実列名.作品ID, 正規化作品ID) || changed;
    値.作品ID = 正規化作品ID;
  }

  if (実列名.形態 && 値.形態) {
    changed = s(実列名.形態, 値.形態) || changed;
  }

  const タイトル = 台湾書籍系_タイトルを生成_(sh.getName(), 値, 設定);
  if (タイトル) {
    changed = s(実列名.タイトル, タイトル) || changed;
  }

  const 重複キー = 台湾書籍系_重複チェックキーを生成_(値);
  if (重複キー && 実列名.重複チェックキー) {
    changed = s(実列名.重複チェックキー, 重複キー) || changed;
  }

  const 不足 = 台湾書籍系_不足項目を返す_(sh.getName(), 値, 設定);

  if (不足.length > 0) {
    changed = s(実列名.コードステータス, '情報不足: ' + 不足.join(',')) || changed;
  } else {
    const codeChanged = 台湾書籍系_親コードSKUを試行生成_(
      sh,
      row,
      設定,
      列,
      rowValues,
      値,
      実列名
    );

    changed = changed || codeChanged;

    const 親コード = 実列名.商品コード
      ? 台湾書籍系_正規化文字列_(rowValues[列[実列名.商品コード] - 1])
      : '';

    const SKU = 実列名.SKU自動
      ? 台湾書籍系_正規化文字列_(rowValues[列[実列名.SKU自動] - 1])
      : '';

    if (親コード || SKU) {
      changed = s(実列名.コードステータス, '生成済み') || changed;
    } else {
      const 現在ステータス = 実列名.コードステータス
        ? 台湾書籍系_正規化文字列_(rowValues[列[実列名.コードステータス] - 1])
        : '';

      if (!現在ステータス || 現在ステータス === '入力途中' || 現在ステータス === '発番待ち') {
        changed = s(実列名.コードステータス, '生成待ち') || changed;
      }
    }
  }

  if (changed) {
    const 粗利益率列 = 実列名.粗利益率 ? 列[実列名.粗利益率] : 0;

    台湾書籍系_1行を書き戻す_粗利益率除外_(
      sh,
      row,
      lastCol,
      rowValues,
      粗利益率列
    );

    if (実列名.作品ID) {
      sh.getRange(row, 列[実列名.作品ID]).setNumberFormat('@');
    }

    if (粗利益率列) {
      sh.getRange(2, 粗利益率列, sh.getMaxRows() - 1, 1).setNumberFormat('0.0%');
    }
  }

  if (works) {
    if (台湾書籍系_正規化文字列_(値.日本語タイトル)) {
      台湾書籍系_Worksのフィールドを更新_(設定, works.作品ID, '日本語タイトル', 値.日本語タイトル);
    }

    if (台湾書籍系_正規化文字列_(値.作者)) {
      台湾書籍系_Worksのフィールドを更新_(設定, works.作品ID, '作者', 値.作者);
    }

    if (台湾書籍系_正規化文字列_(値.原題タイトル)) {
      台湾書籍系_Worksのフィールドを更新_(設定, works.作品ID, '原題タイトル', 値.原題タイトル);
    }

    const 巻数 = 台湾書籍系_数字2桁_(値.単巻数);
    const セット終了 = 台湾書籍系_数字2桁_(値.セット終了);
    const 最新巻候補 = セット終了 ? parseInt(セット終了, 10) : (巻数 ? parseInt(巻数, 10) : null);

    if (最新巻候補 != null) {
      台湾書籍系_Worksの予約巻数を更新_(設定, works.作品ID, 最新巻候補);
    }
  }
}

function 台湾書籍系_WorksIDから取得_(設定, 作品ID) {
  const id = 台湾書籍系_作品ID4桁を取得_(作品ID);
  if (!id) return null;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return null;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  if (col['作品ID'] == null) return null;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (const row of data) {
    const wId = 台湾書籍系_作品ID4桁を取得_(row[col['作品ID']]);
    if (wId === id) {
      return {
        作品ID: wId,
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者: col['作者'] != null ? row[col['作者']] : '',
        原題タイトル: col['原題タイトル'] != null ? row[col['原題タイトル']] : ''
      };
    }
  }
  return null;
}

function 台湾書籍系_現在行とWorksが同じ作品_(値, works) {
  const rowOriginal = 台湾書籍系_作品比較キー_(値.原題タイトル);
  const worksOriginal = 台湾書籍系_作品比較キー_(works.原題タイトル);

  if (rowOriginal && worksOriginal) {
    return rowOriginal === worksOriginal;
  }

  const rowJp = 台湾書籍系_正規化文字列_(値.日本語タイトル);
  const worksJp = 台湾書籍系_正規化文字列_(works.日本語タイトル);

  const rowAuthor = 台湾書籍系_正規化文字列_(値.作者);
  const worksAuthor = 台湾書籍系_正規化文字列_(works.作者);

  if (rowJp && worksJp && rowAuthor && worksAuthor) {
    return rowJp === worksJp && rowAuthor === worksAuthor;
  }

  if (rowJp && worksJp) {
    return rowJp === worksJp;
  }

  // 情報不足の場合は勝手に別作品扱いしない
  return true;
}
function 台湾書籍系_Worksの予約巻数を更新_(設定, 作品ID, 巻数) {
  const id = 台湾書籍系_作品ID4桁を取得_(作品ID);
  if (!id) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  if (col['作品ID'] == null || col['最新巻(予約込み)'] == null) return;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (台湾書籍系_作品ID4桁を取得_(data[i][col['作品ID']]) === id) {
      const 既存 = parseInt(String(data[i][col['最新巻(予約込み)']] || '0'), 10) || 0;
      if (巻数 > 既存) {
        sh.getRange(i + 2, col['最新巻(予約込み)'] + 1).setValue(巻数);
        if (col['予約更新日時'] != null) {
          sh.getRange(i + 2, col['予約更新日時'] + 1).setValue(new Date());
        }
      }
      return;
    }
  }
}
/* ============================================================
 * onEdit 共通
 * ============================================================ */
function 台湾書籍系_onEdit_共通_(e, 設定) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== 設定.マスターシート名) return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  if (開始行 + 行数 - 1 < 2) return;

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const 編集開始列 = e.range.getColumn();
  const 編集終了列 = e.range.getLastColumn();

  const 登録状況列 = 列['登録状況'];
  const 発番発行列 = 列['発番発行'];

  /*
   * 登録状況・発番発行は操作用。
   * ここを触っただけでは補完・採番・Works作成しない。
   *
   * ただし、タイトル・作者・原題・巻数など実データを編集した場合は、
   * onEditからもWorksを取得または作成する。
   */
  const 編集列数 = 編集終了列 - 編集開始列 + 1;

  const 登録状況だけ編集 =
    編集列数 === 1 &&
    登録状況列 &&
    編集開始列 === 登録状況列;

  const 発番発行だけ編集 =
    編集列数 === 1 &&
    発番発行列 &&
    編集開始列 === 発番発行列;

  if (登録状況だけ編集 || 発番発行だけ編集) {
    return;
  }

  const 監視列番号 = (設定.監視列 || [])
    .concat([
      '売価',
      '原価',
      'サイト商品コード',
      '博客來商品コード',
      '作者',
      '日本語タイトル',
      '原題タイトル',
      '原題商品タイトル',
      '言語',
      'カテゴリ',
      '形態(通常/初回限定/特装)',
      '形態（通常/初回限定/特装）',
      '単巻数',
      'セット巻数開始番号',
      'セット巻数終了番号',
      '特典メモ',
      'ISBN'
    ])
    .map(name => 列[台湾書籍系_正規化文字列_(name)])
    .filter(Boolean);

  const 対象列あり = 監視列番号.some(c =>
    c >= 編集開始列 && c <= 編集終了列
  );

  デバッグログ出力_('書籍系_onEdit_共通_到達', {
    sheet: sh.getName(),
    開始行,
    行数,
    編集開始列,
    編集終了列,
    対象列あり,
    Works新規作成: true
  });

  if (!対象列あり) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    デバッグログ出力_('書籍系_onEdit_共通_ロック失敗', {
      sheet: sh.getName(),
      開始行,
      行数
    });
    return;
  }

  try {
    for (let row = 開始行; row < 開始行 + 行数; row++) {
      if (row < 2) continue;

      台湾書籍系_1行補完_共通_(sh, row, 設定, {
        Works新規作成: true
      });
    }
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍系_追加行補完_共通_(sh, startRow, numRows, 設定) {
  if (!sh || numRows <= 0) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r < 2) continue;

      台湾書籍系_1行補完_共通_(sh, r, 設定, {
        Works新規作成: true
      });
    }
  } finally {
    lock.releaseLock();
  }
}
/* ============================================================
 * シート別ラッパー
 * ============================================================ */
/* ============================================================
 * 台湾書籍系 onEdit / 取込後補完 ラッパー（新ルート版）
 * 旧: 台湾書籍系_onEdit_共通_ / 台湾書籍系_追加行補完_共通_
 * 新: _kyoutuu.onEdit処理を実行(...)
 * ============================================================ */

function 台湾書籍系_新onEdit_共通_(e, 設定) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== 設定.マスターシート名) return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();

  if (開始行 + 行数 - 1 < 2) return;

  const 列 = 台湾書籍系_列マップを取得_(sh);

  const 編集開始列 = e.range.getColumn();
  const 編集終了列 = e.range.getLastColumn();

  const 登録状況列 = 列['登録状況'];
  const 発番発行列 = 列['発番発行'];

  /*
   * 登録状況・発番発行は操作用。
   * ここを触っただけでは補完・採番・Works作成しない。
   *
   * タイトル・作者・原題・巻数など実データを編集した場合は、
   * onEditからもWorksを取得または作成する。
   */
  const 編集列数 = 編集終了列 - 編集開始列 + 1;

  const 登録状況だけ編集 =
    編集列数 === 1 &&
    登録状況列 &&
    編集開始列 === 登録状況列;

  const 発番発行だけ編集 =
    編集列数 === 1 &&
    発番発行列 &&
    編集開始列 === 発番発行列;

  if (登録状況だけ編集 || 発番発行だけ編集) {
    return;
  }

  const 監視列名 = [
  ...(設定.監視列 || []),

  '作者',
  '日本語タイトル',
  '原題タイトル',
  '原題商品タイトル',
  'タイトル',
  '言語',
  'カテゴリ',
  '形態',
  '形態(通常/初回限定/特装)',
  '形態（通常/初回限定/特装）',
  '単巻数',
  'セット巻数開始番号',
  'セット巻数終了番号',
  '特典メモ',
  'ISBN',
  '発売日',
  '売価',
  '原価',
  'サイト商品コード',
  '博客來商品コード'
];

const 監視列番号 = [...new Set(監視列名)]
  .map(name => 列[台湾書籍系_正規化文字列_(name)])
  .filter(Boolean);

  const 対象列あり = 監視列番号.some(col =>
    col >= 編集開始列 && col <= 編集終了列
  );

  デバッグログ出力_('書籍系_新onEdit_共通_到達', {
    sheet: sh.getName(),
    開始行,
    行数,
    編集開始列,
    編集終了列,
    対象列あり,
    Works新規作成: true
  });

  if (!対象列あり) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;

  // ループ前に数式を確定し、必要な情報を一括取得
  SpreadsheetApp.flush();

  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(開始行, 1, 行数, lastCol).getValues();

  try {
    for (let r = 開始行; r < 開始行 + 行数; r++) {
      if (r < 2) continue;

      const rowValues = allRowValues[r - 開始行];
      if (!rowValues) continue;

      const 登録状況値 = 登録状況列
        ? String(rowValues[登録状況列 - 1] || '').trim()
        : '';

      // 登録済み行は、通常のonEditでは自動補完・再採番しない
      if (登録状況値.startsWith('登録済')) {
        continue;
      }

      台湾書籍系_1行補完_共通_(sh, r, 設定, {
        Works新規作成: true,
        skipFlush: true,
        列マップ: 列,
        lastCol: lastCol,
        rowValues: rowValues
      });
    }
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍系_新取込後補完_共通_(sheetName, startRow, numRows, cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  if (!sh || !numRows || numRows <= 0) return;

  const 開始行 = Math.max(2, startRow);
  const 終了行 = Math.min(sh.getLastRow(), startRow + numRows - 1);

  if (終了行 < 開始行) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  // ループ前に数式を確定し、必要な情報を一括取得
  SpreadsheetApp.flush();

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(開始行, 1, 終了行 - 開始行 + 1, lastCol).getValues();

  try {
    for (let r = 開始行; r <= 終了行; r++) {
      const rowValues = allRowValues[r - 開始行];
      if (!rowValues) continue;

      台湾書籍系_1行補完_共通_(sh, r, cfg, {
        Works新規作成: true,
        skipFlush: true,
        列マップ: 列,
        lastCol: lastCol,
        rowValues: rowValues
      });
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * シート別ラッパー
 * ============================================================ */

function 台湾まんが_onEdit(e) {
  台湾書籍系_新onEdit_共通_(e, 設定_台湾まんが);
}

function 台湾書籍その他_onEdit(e) {
  台湾書籍系_新onEdit_共通_(e, 設定_台湾書籍その他);
}

function 台湾まんが_取込後補完_(startRow, numRows) {
  台湾書籍系_新取込後補完_共通_('台湾まんが', startRow, numRows, 設定_台湾まんが);
}

function 台湾書籍その他_取込後補完_(startRow, numRows) {
  台湾書籍系_新取込後補完_共通_('台湾書籍その他', startRow, numRows, 設定_台湾書籍その他);
}


function 台湾書籍系_ローカル一括更新_(シート名, 設定) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);

  if (!sh || sh.getLastRow() < 2) {
    ui.alert(`${シート名} にデータがありません`);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const 登録状況列 = headers.indexOf('登録状況') + 1;

  if (!登録状況列) {
    ui.alert('登録状況列が見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const statuses = sh.getRange(2, 登録状況列, lastRow - 1, 1).getDisplayValues();

  const 対象行 = [];
  statuses.forEach((r, i) => {
    const 状態 = String(r[0] || '').trim();
    if (状態 === '未登録') {
      対象行.push(i + 2);
    }
  });

  if (対象行.length === 0) {
    ui.alert('未登録の対象行がありません');
    return;
  }

  if (
    ui.alert(
      '確認',
      `${シート名} の未登録 ${対象行.length} 行をローカル再計算します。\n\n続行しますか？`,
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    対象行.forEach(row => {
  デバッグログ出力_('一括更新_対象行', {
    シート名,
    row,
    Works新規作成: true
  });

  台湾書籍系_1行補完_共通_(sh, row, 設定, {
    Works新規作成: true
  });
});
    ss.toast(`✅ ${シート名}：未登録 ${対象行.length} 行を再計算しました`, '完了', 5);
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍系_Worksを取得または作成_(設定, 値) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定.作品シート名);

  if (!sh) {
    sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品列数).setValues([設定.作品ヘッダー]);
  }

  // Worksの作品ID列は常に文字列固定
  // これをしないと 0087 が 87 に落ちる
  if (sh.getMaxRows() >= 2) {
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }

  const 日本語タイトル = 台湾書籍系_正規化文字列_(値.日本語タイトル);
  const 作者 = 台湾書籍系_正規化文字列_(値.作者);
  const 原題タイトル = 台湾書籍系_正規化文字列_(値.原題タイトル);
  const 既存行ID = 台湾書籍系_行から既存作品IDを取得_(値);

  デバッグログ出力_('Works取得または作成_入力値', {
    日本語タイトル,
    作者,
    原題タイトル,
    既存行ID,
    元_作品ID: 値.作品ID,
    元_商品コード: 値.商品コード,
    元_SKU自動: 値.SKU自動
  });

  // ★商品コードやSKUに既存IDがあるなら、それを最優先でWorksから探す
  if (既存行ID) {
    const byId = 台湾書籍系_WorksIDから取得_(設定, 既存行ID);
    if (byId) {
      return {
        作品ID: 既存行ID,
        日本語タイトル: byId.日本語タイトル || 日本語タイトル,
        作者: byId.作者 || 作者,
        原題タイトル: byId.原題タイトル || 原題タイトル
      };
    }
  }
デバッグログ出力_('Works取得または作成_入力値', {
  日本語タイトル,
  作者,
  原題タイトル,
  元_日本語タイトル: 値.日本語タイトル,
  元_作者: 値.作者,
  元_原題タイトル: 値.原題タイトル
});

  /*
   * まず既存Worksを探す。
   *
   * 例：
   * 作品ID：0007
   * 日本語タイトル：雪雲の花
   * 作者：Snob
   * 原題タイトル：劍鬼花
   *
   * が既にWorksにあれば、それを返す。
   */

const found = 台湾書籍系_Worksから取得_(
  設定,
  原題タイトル,
  日本語タイトル,
  作者
);



if (found && 台湾書籍系_作品ID4桁を取得_(found.作品ID)) {
  return {
    作品ID: 台湾書籍系_作品ID4桁を取得_(found.作品ID),
    日本語タイトル: found.日本語タイトル || 日本語タイトル,
    作者: found.作者 || 作者,
    原題タイトル: found.原題タイトル || 原題タイトル
  };
}

if (!日本語タイトル || !作者) {
  デバッグログ出力_('Works新規作成しない_必須不足', {
    日本語タイトル,
    作者,
    原題タイトル
  });
  return null;
}
  /*
   * 新規IDは、Worksだけでなく商品シート側の使用済みIDも見て採番する。
   * 既存IDは永久IDなので、削除・統合・振り直しはしない。
   */
    // ★既存の商品コード・SKUからIDが取れるなら、それを使う
  // それが無いときだけ、空いている最小番号を採番する
  const 新作品ID = 既存行ID || 台湾書籍系_次の未使用作品ID_();


  const worksKeyBase = 原題タイトル
    ? (台湾書籍系_作品比較キー_(原題タイトル) || 原題タイトル)
    : `${日本語タイトル}||${作者}`;

  const worksKey = 原題タイトル
    ? `原題||${worksKeyBase}`
    : worksKeyBase;

  const writeRow = Math.max(sh.getLastRow() + 1, 2);

  // ★ここに入れる
  デバッグログ出力_('Works新規作成_実行', {
    新作品ID,
    worksKey,
    日本語タイトル,
    作者,
    原題タイトル,
    writeRow
  });

  // 作品ID列は、書き込み直前にも文字列固定
  sh.getRange(writeRow, 2).setNumberFormat('@');

  sh.getRange(writeRow, 1, 1, 設定.作品列数).setValues([[
    worksKey,
    新作品ID,
    日本語タイトル,
    作者,
    原題タイトル,
    '',
    '',
    '',
    '',
    ''
  ]]);

  デバッグログ出力_('Works新規作成_書込完了', {
    新作品ID,
    writeRow
  });

    return {
    作品ID: 新作品ID,
    日本語タイトル,
    作者,
    原題タイトル
  };
}

function デバッグログ出力_(tag, obj) {
  const DEBUG_TRACE = false; // 追跡中だけ true。終わったら false にする。

  if (!DEBUG_TRACE) return;

  try {
    const json = obj ? JSON.stringify(obj) : '';
    console.log(`[${tag}] ${json}`);
  } catch (err) {
    console.log(`[${tag}] ログ出力エラー: ${err}`);
  }
}

function 台湾書籍系_使用済み作品IDセット_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const used = new Set();

  const 正規ID = (v) => {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';

    // 作品ID列用：0007 / 7 / 87 などを4桁化
    if (/^\d+$/.test(s)) {
      return String(parseInt(s, 10)).padStart(4, '0');
    }

    return '';
  };

  const コードからID = (v) => {
    const s = String(v == null ? '' : v).trim().toUpperCase();
    if (!s) return '';

    // 例：
    // TW0088-CM-01
    // TWS0088-CM-0102
    // TWF0088-CM-04
    const m = s.match(/^[A-Z]{2}[A-Z]*?(\d{4})-/);
    return m ? m[1] : '';
  };

  const collectDirectId = (シート名, 候補列名) => {
    const sh = ss.getSheetByName(シート名);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());

    候補列名.forEach(name => {
      const col = headers.indexOf(name) + 1;
      if (!col) return;

      const values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues();

      values.forEach(([v]) => {
        const id = 正規ID(v);
        if (id) used.add(id);
      });
    });
  };

  const collectCodeId = (シート名, 候補列名) => {
    const sh = ss.getSheetByName(シート名);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());

    候補列名.forEach(name => {
      const col = headers.indexOf(name) + 1;
      if (!col) return;

      const values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues();

      values.forEach(([v]) => {
        const id = コードからID(v);
        if (id) used.add(id);
      });
    });
  };

  // Works側：作品ID列だけを見る。空白は未使用扱い。
  collectDirectId('Works（書籍専用）', [
    '作品ID'
  ]);

  // 台湾まんが：作品ID列
  collectDirectId('台湾まんが', [
    '作品ID(W)（自動）',
    '作品ID(W)(自動)',
    '作品(W)（自動）'
  ]);

  // 台湾まんが：コード列
  collectCodeId('台湾まんが', [
    '親コード',
    '商品コード(SKU)',
    '商品コード（SKU）',
    'SKU（自動）',
    'SKU(自動)'
  ]);

  // 台湾書籍その他：作品ID列
  collectDirectId('台湾書籍その他', [
    '作品ID(W)（自動）',
    '作品ID(W)(自動)',
    '作品(W)（自動）'
  ]);

  // 台湾書籍その他：コード列
  collectCodeId('台湾書籍その他', [
    '親コード',
    '商品コード(SKU)',
    '商品コード（SKU）',
    'SKU（自動）',
    'SKU(自動)'
  ]);

  return used;
}

function 台湾書籍系_次の未使用作品ID_() {
  const used = 台湾書籍系_使用済み作品IDセット_();

  // 0001から順に、本当に空いている最小IDを使う
  for (let n = 1; n < 10000; n++) {
    const id = String(n).padStart(4, '0');

    if (!used.has(id)) {
      return id;
    }
  }

  throw new Error('使用可能な作品IDがありません');
}

function 台湾書籍系_作品ID衝突チェック_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = [];

  const addRows = (シート名, id列候補, title列候補, author列候補, code列候補) => {
    const sh = ss.getSheetByName(シート名);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());

    const findCol = (names) => {
      for (const name of names) {
        const idx = headers.indexOf(name);
        if (idx >= 0) return idx + 1;
      }
      return 0;
    };

    const idCol = findCol(id列候補);
    const titleCol = findCol(title列候補);
    const authorCol = findCol(author列候補);
    const codeCol = findCol(code列候補);

    const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();

    values.forEach((r, i) => {
      const idRaw = idCol ? r[idCol - 1] : '';
      const id = 台湾書籍系_作品ID4桁を取得_(idRaw);
      if (!id) return;

      rows.push({
        シート名,
        行: i + 2,
        作品ID: id,
        タイトル: titleCol ? r[titleCol - 1] : '',
        作者: authorCol ? r[authorCol - 1] : '',
        コード: codeCol ? r[codeCol - 1] : ''
      });
    });
  };

  addRows(
    'Works（書籍専用）',
    ['作品ID'],
    ['日本語タイトル', '原題タイトル'],
    ['作者'],
    ['WorksKey']
  );

  addRows(
    '台湾まんが',
    ['作品ID(W)（自動）', '作品ID(W)(自動)', '作品(W)（自動）'],
    ['日本語タイトル'],
    ['作者'],
    ['親コード', '商品コード(SKU)', 'SKU（自動）', 'SKU(自動)']
  );

  addRows(
    '台湾書籍その他',
    ['作品ID(W)（自動）', '作品ID(W)(自動)', '作品(W)（自動）'],
    ['日本語タイトル'],
    ['作者'],
    ['親コード', '商品コード(SKU)', 'SKU（自動）', 'SKU(自動)']
  );

  const byId = {};
  rows.forEach(r => {
    if (!byId[r.作品ID]) byId[r.作品ID] = [];
    byId[r.作品ID].push(r);
  });

  Object.keys(byId).sort().forEach(id => {
    const list = byId[id];

    const titles = [...new Set(list.map(x => `${x.タイトル} / ${x.作者}`).filter(Boolean))];

    if (titles.length >= 2) {
      console.log('⚠️ 作品ID衝突: ' + id);
      list.forEach(x => console.log(JSON.stringify(x)));
    }
  });

  SpreadsheetApp.getUi().alert('作品ID衝突チェックを実行しました。Apps Script の「実行数」ログを確認してください。');
}

function 台湾書籍系_Works作品IDを4桁文字列に統一_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Works（書籍専用）');

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Works（書籍専用）にデータがありません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const idCol = headers.indexOf('作品ID') + 1;

  if (!idCol) {
    SpreadsheetApp.getUi().alert('Worksに「作品ID」列が見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const range = sh.getRange(2, idCol, lastRow - 1, 1);
  const values = range.getDisplayValues();

  let 更新数 = 0;

  const out = values.map(([v]) => {
    const s = String(v || '').trim();

    if (!s) return [''];

    const m = s.match(/\d+/);
    if (!m) return [s];

    const id = String(parseInt(m[0], 10)).padStart(4, '0');

    if (s !== id) 更新数++;

    return [id];
  });

  // 先に文字列書式にしてから値を入れる
  range.setNumberFormat('@');
  range.setValues(out);

  SpreadsheetApp.getUi().alert(
    `✅ Worksの作品IDを4桁文字列に統一しました\n\n更新: ${更新数}件`
  );
}

function 作品ID_次の空き番号を確認() {
  const used = 台湾書籍系_使用済み作品IDセット_();

  const 空き = [];
  for (let n = 1; n < 10000; n++) {
    const id = String(n).padStart(4, '0');
    if (!used.has(id)) 空き.push(id);
    if (空き.length >= 20) break;
  }

  console.log('使用済みID件数: ' + used.size);
  console.log('次の空きID候補: ' + 空き.join(', '));

  SpreadsheetApp.getUi().alert(
    '次に使う作品ID: ' + (空き[0] || 'なし') +
    '\n\n空き候補:\n' + 空き.join(', ')
  );
}

function 台湾書籍系_コードから作品ID4桁を取得_(v) {
  const s = 台湾書籍系_正規化文字列_(v).toUpperCase();
  if (!s) return '';

  // 例:
  // TW0002-CM-16
  // TWF0002-CM-17
  // TWS0007-CM-0304
  const m = s.match(/^[A-Z]{2}[A-Z]*?(\d{4})-/);
  return m ? m[1] : '';
}

function 台湾書籍系_行から既存作品IDを取得_(値) {
  return (
    台湾書籍系_作品ID4桁を取得_(値.作品ID) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.商品コード) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.SKU自動) ||
    台湾書籍系_コードから作品ID4桁を取得_(値.親コード) ||
    ''
  );
}

function 台湾書籍系_チェック行を事前補完_(シート名, 設定) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);
  if (!sh || sh.getLastRow() < 2) return;

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const チェック列 = 列['発番発行'];
  if (!チェック列) return;

  const lastRow = sh.getLastRow();
  const checks = sh.getRange(2, チェック列, lastRow - 1, 1).getValues();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    checks.forEach((r, i) => {
      if (r[0] === true) {
        const row = i + 2;
        台湾書籍系_1行補完_共通_(sh, row, 設定, {
          Works新規作成: true
        });
      }
    });

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍系_Worksのフィールドを更新_(設定, 作品ID, フィールド名, 新値) {
  const id = 台湾書籍系_作品ID4桁を取得_(作品ID);
  if (!id || !新値) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  if (col['作品ID'] == null || col[フィールド名] == null) return;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (台湾書籍系_作品ID4桁を取得_(data[i][col['作品ID']]) === id) {
      const 現在値 = 台湾書籍系_正規化文字列_(data[i][col[フィールド名]]);
      const 新値正規化 = 台湾書籍系_正規化文字列_(新値);
      if (現在値 !== 新値正規化) {
        sh.getRange(i + 2, col[フィールド名] + 1).setValue(新値);
      }
      return;
    }
  }
}

function 台湾書籍系_1行を書き戻す_粗利益率除外_(sh, row, lastCol, rowValues, 粗利益率列) {
  // 粗利益率列が見つからない場合は従来通り
  if (!粗利益率列) {
    sh.getRange(row, 1, 1, lastCol).setValues([rowValues]);
    return;
  }

  // 粗利益率列の左側だけ書き込み
  if (粗利益率列 > 1) {
    sh.getRange(row, 1, 1, 粗利益率列 - 1)
      .setValues([rowValues.slice(0, 粗利益率列 - 1)]);
  }

  // 粗利益率列の右側だけ書き込み
  if (粗利益率列 < lastCol) {
    sh.getRange(row, 粗利益率列 + 1, 1, lastCol - 粗利益率列)
      .setValues([rowValues.slice(粗利益率列)]);
  }

  // N2などARRAYFORMULA本体は消さない。
  // 3行目以降は静的値が残っていたら消して、ARRAYFORMULAが展開できるようにする。
  if (row > 2) {
    sh.getRange(row, 粗利益率列).clearContent();
  }
}

function 台湾まんが_タイトルの形態括弧だけ高速修正() {
  台湾書籍系_タイトル形態括弧だけ高速修正_('台湾まんが');
}

function 台湾書籍その他_タイトルの形態括弧だけ高速修正() {
  台湾書籍系_タイトル形態括弧だけ高速修正_('台湾書籍その他');
}

function 台湾書籍系_タイトル形態括弧だけ高速修正_(シート名) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(`${シート名} にデータがありません`);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const titleCol = headers.indexOf('タイトル') + 1;
  if (!titleCol) {
    SpreadsheetApp.getUi().alert('タイトル列が見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const range = sh.getRange(2, titleCol, lastRow - 1, 1);
  const values = range.getDisplayValues();

  const 形態一覧 = [
    '予約限定特装版',
    '初版限定版',
    '初回限定版',
    '首刷限定版',
    '特装版',
    '特裝版',
    '限定版'
  ];

  const re = new RegExp(
    `^(.*?\\s+まんが)\\s+(${形態一覧.join('|')})(\\s+『)`,
    'u'
  );

  let changed = 0;

  for (let i = 0; i < values.length; i++) {
    const before = String(values[i][0] || '');
    if (!before) continue;

    const after = before.replace(re, '$1($2)$3');

    if (after !== before) {
      values[i][0] = after;
      changed++;
    }
  }

  if (changed > 0) {
    range.setValues(values);
  }

  SpreadsheetApp.getUi().alert(`完了：タイトル ${changed}件を修正しました`);
}


function 台湾まんが_登録済み行のステータスを確定に戻す_高速() {
  台湾書籍系_登録済み行のステータスを確定に戻す_高速_('台湾まんが');
}

function 台湾書籍その他_登録済み行のステータスを確定に戻す_高速() {
  台湾書籍系_登録済み行のステータスを確定に戻す_高速_('台湾書籍その他');
}

function 台湾書籍系_登録済み行のステータスを確定に戻す_高速_(シート名) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);

  if (!sh || sh.getLastRow() < 2) {
    ss.toast(`${シート名} に対象データがありません`, '完了', 5);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const col登録状況 = headers.indexOf('登録状況') + 1;
  const colステータス = headers.indexOf('商品コードステータス') + 1;

  if (!col登録状況 || !colステータス) {
    SpreadsheetApp.getUi().alert('登録状況列または商品コードステータス列が見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const rowCount = lastRow - 1;
  const startRow = 2;

  const 確定ステータス = '商品コード(発行済み確定)';

  // 読み込みは2列だけ
  const 登録状況 = sh.getRange(startRow, col登録状況, rowCount, 1).getDisplayValues();
  const 現在ステータス = sh.getRange(startRow, colステータス, rowCount, 1).getDisplayValues();

  const 更新ブロック = [];
  let blockStartRow = null;
  let blockLength = 0;
  let changed = 0;

  for (let i = 0; i < rowCount; i++) {
    const reg = String(登録状況[i][0] || '').trim();
    const status = String(現在ステータス[i][0] || '').trim();
    const actualRow = startRow + i;

    const 要更新 =
      reg.startsWith('登録済') &&
      status !== 確定ステータス;

    if (!要更新) {
      if (blockStartRow !== null) {
        更新ブロック.push({
          startRow: blockStartRow,
          length: blockLength
        });
        blockStartRow = null;
        blockLength = 0;
      }
      continue;
    }

    changed++;

    if (blockStartRow === null) {
      blockStartRow = actualRow;
      blockLength = 1;
    } else {
      blockLength++;
    }
  }

  if (blockStartRow !== null) {
    更新ブロック.push({
      startRow: blockStartRow,
      length: blockLength
    });
  }

  // 変更が必要な連続ブロックだけ書き込み
  更新ブロック.forEach(block => {
    const values = Array.from(
      { length: block.length },
      () => [確定ステータス]
    );

    sh.getRange(block.startRow, colステータス, block.length, 1)
      .setValues(values);
  });

  ss.toast(
    `完了：${changed}件のステータスを確定に戻しました`,
    '完了',
    5
  );
}

function 台湾まんが_登録済み作品IDを商品コードから復元_超高速() {
  台湾書籍系_登録済み作品IDを商品コードから復元_超高速_('台湾まんが');
}

function 台湾書籍その他_登録済み作品IDを商品コードから復元_超高速() {
  台湾書籍系_登録済み作品IDを商品コードから復元_超高速_('台湾書籍その他');
}

function 台湾書籍系_列番号候補取得_(headers, 候補) {
  const normalizedHeaders = headers.map(h => 台湾書籍系_正規化文字列_(h));

  for (const name of 候補) {
    const key = 台湾書籍系_正規化文字列_(name);
    const idx = normalizedHeaders.indexOf(key);
    if (idx >= 0) return idx + 1;
  }

  return 0;
}

function 台湾書籍系_登録済み作品IDを商品コードから復元_超高速_(シート名) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);

  if (!sh || sh.getLastRow() < 2) {
    ss.toast(`${シート名} に対象データがありません`, '完了', 5);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const col登録状況 = 台湾書籍系_列番号候補取得_(headers, [
    '登録状況'
  ]);

  const col作品ID = 台湾書籍系_列番号候補取得_(headers, [
    '作品ID(W)（自動）',
    '作品ID(W)(自動)',
    '作品(W)（自動）'
  ]);

  const col商品コード = 台湾書籍系_列番号候補取得_(headers, [
    '親コード',
    '商品コード',
    '商品コード(SKU)',
    '商品コード（SKU）'
  ]);

  const colSKU = 台湾書籍系_列番号候補取得_(headers, [
    'SKU（自動）',
    'SKU(自動)'
  ]);

  const colステータス = 台湾書籍系_列番号候補取得_(headers, [
    '商品コードステータス'
  ]);

  if (!col登録状況 || !col作品ID || (!col商品コード && !colSKU)) {
    SpreadsheetApp.getUi().alert(
      '登録状況 / 作品ID / 商品コードまたはSKU列が見つかりません'
    );
    return;
  }

  const startRow = 2;
  const rowCount = sh.getLastRow() - 1;

  const 登録状況 = sh.getRange(startRow, col登録状況, rowCount, 1).getDisplayValues();
  const 作品ID = sh.getRange(startRow, col作品ID, rowCount, 1).getDisplayValues();

  const 商品コード = col商品コード
    ? sh.getRange(startRow, col商品コード, rowCount, 1).getDisplayValues()
    : Array.from({ length: rowCount }, () => ['']);

  const SKU = colSKU
    ? sh.getRange(startRow, colSKU, rowCount, 1).getDisplayValues()
    : Array.from({ length: rowCount }, () => ['']);

  const ステータス = colステータス
    ? sh.getRange(startRow, colステータス, rowCount, 1).getDisplayValues()
    : null;

  const fixedStatus = '商品コード(発行済み確定)';

  let changedId = 0;
  let changedStatus = 0;

  // 配列上でまとめて修正
  for (let i = 0; i < rowCount; i++) {
    const reg = String(登録状況[i][0] || '').trim();
    if (!reg.startsWith('登録済')) continue;

    const code = String(商品コード[i][0] || '').trim();
    const sku = String(SKU[i][0] || '').trim();

    const idFromCode =
      台湾書籍系_コードから作品ID4桁を取得_(code) ||
      台湾書籍系_コードから作品ID4桁を取得_(sku);

    if (idFromCode) {
      const currentId = 台湾書籍系_作品ID4桁を取得_(作品ID[i][0]);

      if (currentId !== idFromCode) {
        作品ID[i][0] = idFromCode;
        changedId++;
      }
    }

    if (ステータス) {
      const currentStatus = String(ステータス[i][0] || '').trim();

      if (currentStatus !== fixedStatus) {
        ステータス[i][0] = fixedStatus;
        changedStatus++;
      }
    }
  }

  // 書き込みは最大2回だけ
  if (changedId > 0) {
    sh.getRange(startRow, col作品ID, rowCount, 1).setValues(作品ID);
  }

  if (colステータス && ステータス && changedStatus > 0) {
    sh.getRange(startRow, colステータス, rowCount, 1).setValues(ステータス);
  }

  ss.toast(
    `完了：作品ID ${changedId}件 / ステータス ${changedStatus}件を修復しました`,
    '完了',
    7
  );
}

function 修復_枯れた花に涙_Works作品IDを0102に戻す() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Works（書籍専用）');
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Works（書籍専用）シートが見つからないか、データがありません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const col作品ID = headers.indexOf('作品ID') + 1;
  const col日本語タイトル = headers.indexOf('日本語タイトル') + 1;
  const col原題タイトル = headers.indexOf('原題タイトル') + 1;

  if (!col作品ID || !col日本語タイトル || !col原題タイトル) {
    SpreadsheetApp.getUi().alert('Worksの 作品ID / 日本語タイトル / 原題タイトル 列が見つかりません');
    return;
  }

  const rowCount = sh.getLastRow() - 1;
  const values = sh.getRange(2, 1, rowCount, sh.getLastColumn()).getValues();

  let changed = 0;

  for (let i = 0; i < values.length; i++) {
    const jp = String(values[i][col日本語タイトル - 1] || '').trim();
    const original = String(values[i][col原題タイトル - 1] || '').trim();

    const isTarget =
      jp.includes('枯れた花に涙') ||
      original.includes('枯れた花に涙') ||
      original.includes('枯萎的花淚');

    if (!isTarget) continue;

    if (String(values[i][col作品ID - 1]).trim() !== '0102') {
      values[i][col作品ID - 1] = '0102';
      changed++;
    }
  }

  if (changed > 0) {
    sh.getRange(2, 1, rowCount, sh.getLastColumn()).setValues(values);
    sh.getRange(2, col作品ID, rowCount, 1).setNumberFormat('@');
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `完了：Worksの作品IDを ${changed}件 修復しました`,
    '修復完了',
    5
  );
}

function 台湾まんが_登録済みタイトルの形態カッコだけ修正() {
  台湾書籍系_登録済みタイトルの形態カッコだけ修正_('台湾まんが');
}

function 台湾書籍系_登録済みタイトルの形態カッコだけ修正_(シート名) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);
  if (!sh || sh.getLastRow() < 2) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const col登録状況 = headers.indexOf('登録状況') + 1;
  const colタイトル = headers.indexOf('タイトル') + 1;

  if (!col登録状況 || !colタイトル) {
    SpreadsheetApp.getUi().alert('登録状況列またはタイトル列が見つかりません');
    return;
  }

  const startRow = 2;
  const rowCount = sh.getLastRow() - 1;

  const 登録状況 = sh.getRange(startRow, col登録状況, rowCount, 1).getDisplayValues();
  const タイトル = sh.getRange(startRow, colタイトル, rowCount, 1).getDisplayValues();

  const 形態一覧 = [
    '予約限定特装版',
    '予約限定特装',
    '初版限定版',
    '初回限定版',
    '首刷限定版',
    '特装版',
    '特裝版',
    '限定版'
  ];

  const re = new RegExp(
    `(台湾版\\s+まんが)\\s+(${形態一覧.join('|')})(\\s+[『「])`,
    'u'
  );

  let changed = 0;

  for (let i = 0; i < rowCount; i++) {
    const reg = String(登録状況[i][0] || '').trim();
    if (!reg.startsWith('登録済')) continue;

    const before = String(タイトル[i][0] || '');
    if (!before) continue;

    const after = before.replace(re, '$1($2)$3');

    if (after !== before) {
      タイトル[i][0] = after;
      changed++;
    }
  }

  if (changed > 0) {
    sh.getRange(startRow, colタイトル, rowCount, 1).setValues(タイトル);
  }

  SpreadsheetApp.getUi().alert('登録状況列またはタイトル列が見つかりません');
}