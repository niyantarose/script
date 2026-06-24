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

var _台湾書籍系_媒体コード集合キャッシュ = null;

/**
 * WorksKey の媒体セグメント判定に使う「正規な媒体コード集合」。
 * ハードコードの正規表現ではなくカテゴリマスター（カテゴリ→コード）から取得し、
 * カテゴリ追加時も許可リストが自動で広がるようにする。実行内キャッシュ。
 */
function 台湾書籍系_媒体コード集合_() {
  if (_台湾書籍系_媒体コード集合キャッシュ) return _台湾書籍系_媒体コード集合キャッシュ;
  const set = {};
  const 設定 =
    (typeof 設定_台湾まんが !== 'undefined' && 設定_台湾まんが) ||
    (typeof 設定_台湾書籍その他 !== 'undefined' && 設定_台湾書籍その他) ||
    null;
  try {
    const master = 台湾書籍系_カテゴリマスターを取得_(設定);
    const codeMap = (master && master.codeToName) || {};
    Object.keys(codeMap).forEach(code => {
      const c = 台湾書籍系_正規化文字列_(code).toUpperCase();
      if (c) set[c] = true;
    });
  } catch (err) {
    console.log('台湾書籍系_媒体コード集合_ error: ' + err);
  }
  // マスター未読込・読込失敗時の保険。最低限の媒体コードは常に有効にする。
  ['CM', 'NV', 'ART', 'MZ', 'GD', 'ES'].forEach(c => { set[c] = true; });
  _台湾書籍系_媒体コード集合キャッシュ = set;
  return set;
}

function 台湾書籍系_WorksKey媒体コード_(worksKey) {
  const key = 台湾書籍系_正規化文字列_(worksKey);
  if (!key) return '';
  const 媒体集合 = 台湾書籍系_媒体コード集合_();
  // セグメントが「カテゴリマスターに存在する正規コード」のときだけ媒体とみなす。
  // 英語題などの旧キー（例 SPY||作者）を媒体と誤認しないため、形式一致では判定しない。
  const isMedia = seg => {
    const v = 台湾書籍系_正規化文字列_(seg).toUpperCase();
    return !!v && 媒体集合[v] === true;
  };
  const parts = key.split('||').map(v => 台湾書籍系_正規化文字列_(v).toUpperCase());
  if (parts.length >= 3 && parts[0] === '原題' && isMedia(parts[1])) return parts[1];
  if (parts.length >= 2 && isMedia(parts[0])) return parts[0];
  return '';
}

function 台湾書籍系_Works行媒体コード_(row, col) {
  if (!row || !col || col['WorksKey'] == null) return '';
  return 台湾書籍系_WorksKey媒体コード_(row[col['WorksKey']]);
}

function 台湾書籍系_Works行が媒体一致_(row, col, 媒体コード) {
  const target = 台湾書籍系_正規化文字列_(媒体コード).toUpperCase();
  if (!target) return true;
  const rowMedia = 台湾書籍系_Works行媒体コード_(row, col);
  // 媒体なしの旧WorksKeyは後方互換のため一致扱い。移行後は媒体付きキーで厳密一致する。
  if (!rowMedia) return true;
  return rowMedia === target;
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
  if (v === '予約限定特装版' || v === '予約限定特装') return 'RS';
  if (v === '限定特装版' || v === '限定特装') return 'SS';
  if (v === '初版限定版') return 'F';
  if (v === '初回限定版') return 'F';
  if (v === '特装版' || v === '特裝版') return 'S';
  if (v === '限定版') return 'S';

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
  const result = {
    nameToCode: {},
    codeToName: {},
    nameToBonusText: {}
  };

  const sheetNames = [
    設定 && 設定.形態マスター名,
    '形態マスターシート',
    '形態マスター',
    '形態'
  ]
    .map(v => 台湾書籍系_正規化文字列_(v))
    .filter(Boolean);

  let sh = null;
  for (const name of [...new Set(sheetNames)]) {
    sh = ss.getSheetByName(name);
    if (sh) break;
  }

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
function 台湾書籍系_Worksから取得_(設定, 原題タイトル, 日本語タイトル = '', 作者 = '', 媒体コード = '') {
  const key = 台湾書籍系_正規化文字列_(原題タイトル);
  // ★追加
  
  const 作品キー = 台湾書籍系_作品比較キー_(原題タイトル);
  const jpKey = 台湾書籍系_正規化文字列_(日本語タイトル);
  const authorKey = 台湾書籍系_正規化文字列_(作者);
  const mediaKey = 台湾書籍系_正規化文字列_(媒体コード).toUpperCase();

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
    if (!台湾書籍系_Works行が媒体一致_(row, col, mediaKey)) continue;
    const raw = 台湾書籍系_正規化文字列_(row[col['原題タイトル']]);
    if (raw && raw === key) {
      return {
        作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者: col['作者'] != null ? row[col['作者']] : '',
        原題タイトル: col['原題タイトル'] != null ? row[col['原題タイトル']] : ''
      };
    }
  }

  // 2) 原題タイトル 作品比較キー一致
  for (const row of data) {
    if (!台湾書籍系_Works行が媒体一致_(row, col, mediaKey)) continue;
    const raw = 台湾書籍系_正規化文字列_(row[col['原題タイトル']]);
    const raw作品キー = 台湾書籍系_作品比較キー_(raw);
    if (raw作品キー && 作品キー && raw作品キー === 作品キー) {
      return {
        作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者: col['作者'] != null ? row[col['作者']] : '',
        原題タイトル: col['原題タイトル'] != null ? row[col['原題タイトル']] : ''
      };
    }
  }

  // 3) 日本語タイトル + 作者 fallback
  if (jpKey && authorKey && col['日本語タイトル'] != null && col['作者'] != null) {
    for (const row of data) {
      if (!台湾書籍系_Works行が媒体一致_(row, col, mediaKey)) continue;
      const jp = 台湾書籍系_正規化文字列_(row[col['日本語タイトル']]);
      const au = 台湾書籍系_正規化文字列_(row[col['作者']]);
      if (jp === jpKey && au === authorKey) {
        return {
          作品ID: col['作品ID'] != null ? row[col['作品ID']] : '',
          日本語タイトル: jp,
          作者: au,
          原題タイトル: col['原題タイトル'] != null ? row[col['原題タイトル']] : ''
        };
      }
    }
  }

  return null;
}

function 台湾書籍系_作品照合キー候補_(値) {
  const candidates = [
    値 && 値.原題タイトル,
    値 && 値.原題商品タイトル,
    値 && 値.日本語タイトル
  ];
  const keys = [];
  candidates.forEach(v => {
    const key = 台湾書籍系_作品比較キー_(v);
    if (key && !keys.includes(key)) keys.push(key);
  });
  return keys;
}

function 台湾書籍系_Worksから同一作品情報を取得_(設定, 値) {
  const 媒体コード = 台湾書籍系_媒体コードを生成_(値, 設定 && 設定.マスターシート名);
  const candidates = [
    値 && 値.原題タイトル,
    値 && 値.原題商品タイトル,
    値 && 値.日本語タイトル
  ].map(v => 台湾書籍系_正規化文字列_(v)).filter(Boolean);

  for (const title of candidates) {
    const found = 台湾書籍系_Worksから取得_(設定, title, 値.日本語タイトル, 値.作者, 媒体コード);
    if (found && 台湾書籍系_作品ID4桁を取得_(found.作品ID)) {
      return found;
    }
  }
  return null;
}

function 台湾書籍系_同一作品行から情報を取得_(sh, currentRow, 列, 値, 実列名) {
  const targetKeys = 台湾書籍系_作品照合キー候補_(値);
  if (!sh || !targetKeys.length || sh.getLastRow() < 2) return null;
  const targetMedia = 台湾書籍系_媒体コードを生成_(値, sh.getName());

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const read = (row, colName) => {
    if (!colName || !列[colName]) return '';
    return row[列[colName] - 1];
  };

  let best = null;
  let bestScore = -1;

  values.forEach((row, i) => {
    const rowNo = i + 2;
    if (rowNo === currentRow) return;

    const candidate = {
      日本語タイトル: read(row, 実列名.日本語タイトル),
      作者: read(row, 実列名.作者),
      原題タイトル: read(row, 実列名.原題),
      原題商品タイトル: read(row, 実列名.原題商品タイトル),
      言語: read(row, 実列名.言語),
      カテゴリ: read(row, 実列名.カテゴリ),
      作品ID: read(row, 実列名.作品ID),
      商品コード: read(row, 実列名.商品コード),
      親コード: read(row, 実列名.商品コード),
      SKU自動: read(row, 実列名.SKU自動)
    };

    const candidateMedia = 台湾書籍系_媒体コードを生成_(candidate, sh.getName());
    if (targetMedia && candidateMedia && targetMedia !== candidateMedia) return;

    const candidateKeys = 台湾書籍系_作品照合キー候補_(candidate);
    if (!candidateKeys.some(key => targetKeys.includes(key))) return;

    const id = 台湾書籍系_行から既存作品IDを取得_(candidate);
    const score =
      (id ? 100 : 0) +
      (台湾書籍系_正規化文字列_(candidate.作者) ? 20 : 0) +
      (台湾書籍系_正規化文字列_(candidate.日本語タイトル) ? 10 : 0) +
      (台湾書籍系_正規化文字列_(candidate.原題タイトル) ? 10 : 0) +
      (台湾書籍系_正規化文字列_(candidate.言語) ? 3 : 0) +
      (台湾書籍系_正規化文字列_(candidate.カテゴリ) ? 3 : 0);

    if (score > bestScore) {
      bestScore = score;
      best = Object.assign({}, candidate, { 作品ID: id || candidate.作品ID });
    }
  });

  return best;
}

function 台湾書籍系_同一作品情報を補完_(sh, row, 設定, 列, 値, 実列名, setValue) {
  const sheetInfo = 台湾書籍系_同一作品行から情報を取得_(sh, row, 列, 値, 実列名);
  const worksInfo = 台湾書籍系_Worksから同一作品情報を取得_(設定, 値);
  const info = Object.assign({}, sheetInfo || {}, worksInfo || {});

  if (!Object.keys(info).length) return false;

  let changed = false;
  const fill = (prop, colName, v) => {
    const value = 台湾書籍系_正規化文字列_(v);
    if (!value) return;
    if (台湾書籍系_正規化文字列_(値[prop])) return;
    値[prop] = value;
    changed = setValue(colName, value) || changed;
  };

  fill('作品ID', 実列名.作品ID, 台湾書籍系_作品ID4桁を取得_(info.作品ID));
  fill('日本語タイトル', 実列名.日本語タイトル, info.日本語タイトル);
  fill('作者', 実列名.作者, info.作者);
  fill('原題タイトル', 実列名.原題, info.原題タイトル);
  fill('言語', 実列名.言語, info.言語);
  fill('カテゴリ', 実列名.カテゴリ, info.カテゴリ);

  return changed;
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
  // 台湾書籍その他は小説中心。カテゴリ未入力でも NV に寄せる。
  // （空カテゴリで BK に落ちると、同一小説が NV と BK で WorksKey が食い違い、別作品IDが振られるため）
  if (台湾書籍系_正規化文字列_(シート名) === '台湾書籍その他') return 'NV';

  const raw = v.toUpperCase().replace(/[^A-Z0-9]/g, '');
  return raw ? raw.slice(0, 3) : 'BK';
}

function 台湾書籍系_媒体コードを生成_(値, シート名) {
  const カテゴリ = 値 && typeof 値 === 'object' ? 値.カテゴリ : 値;
  return 台湾書籍系_カテゴリコードを生成_(カテゴリ, シート名);
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
      // 「2完」「45完結」のように数字の後に完/完結が来る中華題でも巻数を取れるよう、完結|完 を追加
      // (4) 特裝版 のような括弧＋版種キーワードのパターンに対応するため、lookahead内に optional な括弧を追加
      /(?:^|[^\d])(\d{1,3})(?=\s*[\(\[（]?\s*[\)\]）]?\s*(?:特装版|特裝版|首刷限定版|初版限定版|初回限定版|限定版|通常版|特装|特裝|首刷|初版|版|特典|完結|完|$))/i,
      // 括弧囲みの数字を直接抽出するパターンを追加 (例: (3) 完)
      /[(（](\d{1,3})[)）]/i
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

function 台湾書籍系_商品コードベース化_(code) {
  return 台湾書籍系_正規化文字列_(code).replace(/-(?:\d{2}|\d{4}|SET)$/i, '');
}

function 台湾書籍系_商品コード末尾キー_(code) {
  const base = 台湾書籍系_商品コードベース化_(code).toUpperCase();
  const m = base.match(/(\d{4}-[A-Z0-9]+)$/);
  return m ? m[1] : '';
}

function 台湾書籍系_商品コード構成を取得_(code) {
  const base = 台湾書籍系_商品コードベース化_(code).toUpperCase();
  const m = base.match(/^([A-Z]{2})([A-Z]*)(\d{4})-([A-Z0-9]+)$/);
  if (!m) return null;
  return {
    言語コード: m[1],
    形態コード: m[2] || '',
    作品ID: m[3],
    カテゴリコード: m[4]
  };
}

function 台湾書籍系_既存ベースを優先してよい_(既存コード, 生成コード) {
  const 既存ベース = 台湾書籍系_商品コードベース化_(既存コード);
  if (!既存ベース) return false;
  const 生成ベース = 台湾書籍系_商品コードベース化_(生成コード);
  if (!生成ベース) return true;
  const 既存キー = 台湾書籍系_商品コード末尾キー_(既存ベース);
  const 生成キー = 台湾書籍系_商品コード末尾キー_(生成ベース);
  if (!既存キー || !生成キー || 既存キー !== 生成キー) return false;

  const 既存構成 = 台湾書籍系_商品コード構成を取得_(既存ベース);
  const 生成構成 = 台湾書籍系_商品コード構成を取得_(生成ベース);
  if (!既存構成 || !生成構成) return true;

  if (既存構成.言語コード !== 生成構成.言語コード) return false;

  // 生成側に形態コードがあるなら、その形態を正とする。
  // 生成側が空の時だけ、既存コードの F/S/RS などを保険として残す。
  if (生成構成.形態コード && 既存構成.形態コード !== 生成構成.形態コード) {
    return false;
  }

  return true;
}

function 台湾書籍系_最終SKUを組み立て_(baseCode, 値) {
  const base = 台湾書籍系_商品コードベース化_(baseCode);
  const volumeCode = 台湾書籍系_巻数を抽出_(値);
  if (!base) return '';
  if (!volumeCode) return base;
  return `${base}-${volumeCode}`;
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

  const 既存ベース候補 =
    台湾書籍系_正規化文字列_(値.商品コード) ||
    台湾書籍系_正規化文字列_(値.親コード) ||
    台湾書籍系_正規化文字列_(値.SKU自動);
  if (台湾書籍系_既存ベースを優先してよい_(既存ベース候補, base)) {
    base = 台湾書籍系_商品コードベース化_(既存ベース候補);
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
  if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
    台湾まんが_ONEDIT_LOG_('code:result', {
      sheet: sh.getName(),
      row,
      volume: 値.単巻数,
      setStart: 値.セット開始,
      setEnd: 値.セット終了,
      beforeCode: 現在親コード,
      baseCode: 生成後親コード,
      sku: SKU列名 ? rowValues[列[SKU列名] - 1] : '',
      finalCode: 商品コード列名 ? rowValues[列[商品コード列名] - 1] : '',
      codeColumn: 商品コード列名 ? 列[商品コード列名] : 0,
      skuColumn: SKU列名 ? 列[SKU列名] : 0
    });
  }

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

  const modifiedCols = [];
  const s = (列名, 値) => {
    if (!列名 || !列[列名]) return false;
    const idx = 列[列名] - 1;
    const before = rowValues[idx];
    if (String(before) === String(値)) return false;
    rowValues[idx] = 値;
    if (!modifiedCols.includes(列名)) {
      modifiedCols.push(列名);
    }
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

  changed = 台湾書籍系_同一作品情報を補完_(
    sh,
    row,
    設定,
    列,
    値,
    実列名,
    s
  ) || changed;

  if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
    台湾まんが_ONEDIT_LOG_('row:start', {
      sheet: sh.getName(),
      row,
      codeColumn: 実列名.商品コード ? 列[実列名.商品コード] : 0,
      skuColumn: 実列名.SKU自動 ? 列[実列名.SKU自動] : 0,
      volumeColumn: 実列名.単巻数 ? 列[実列名.単巻数] : 0,
      setStartColumn: 実列名.セット開始 ? 列[実列名.セット開始] : 0,
      setEndColumn: 実列名.セット終了 ? 列[実列名.セット終了] : 0,
      currentCode: 値.商品コード,
      currentSku: 値.SKU自動,
      volume: 値.単巻数,
      setStart: 値.セット開始,
      setEnd: 値.セット終了,
      status: 値.登録状況
    });
  }

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
      modifiedCols.forEach(colName => {
        const colIndex = 列[colName];
        if (colIndex) {
          sh.getRange(row, colIndex).setValue(rowValues[colIndex - 1]);
        }
      });

      const 粗利益率列 = 実列名.粗利益率 ? 列[実列名.粗利益率] : 0;
      if (粗利益率列 && row > 2) {
        sh.getRange(row, 粗利益率列).clearContent();
      }

      if (実列名.作品ID && modifiedCols.includes(実列名.作品ID)) {
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
    : 台湾書籍系_Worksから取得_(設定, 値.原題タイトル, 値.日本語タイトル, 値.作者, 台湾書籍系_媒体コードを生成_(値, sh.getName()));

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

    if (codeChanged) {
      [実列名.商品コード, 実列名.SKU自動].forEach(colName => {
        if (colName && 列[colName] && !modifiedCols.includes(colName)) {
          modifiedCols.push(colName);
        }
      });
    }

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
    modifiedCols.forEach(colName => {
      const colIndex = 列[colName];
      if (colIndex) {
        sh.getRange(row, colIndex).setValue(rowValues[colIndex - 1]);
      }
    });

    const 粗利益率列 = 実列名.粗利益率 ? 列[実列名.粗利益率] : 0;
    if (粗利益率列) {
      // ★高速化: 以前は毎回シート全体（最大数万行）に setNumberFormat していたため非常に遅かった。
      //   生成対象の行だけを %書式にすれば十分（各行は生成時に必ずここを通る）。
      const 粗利益率セル = sh.getRange(row, 粗利益率列);
      if (row > 2) {
        粗利益率セル.clearContent();
      }
      粗利益率セル.setNumberFormat('0.0%');
    }

    if (実列名.作品ID && modifiedCols.includes(実列名.作品ID)) {
      sh.getRange(row, 列[実列名.作品ID]).setNumberFormat('@');
    }
  }
  if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
    台湾まんが_ONEDIT_LOG_('row:write_done', {
      sheet: sh.getName(),
      row,
      changed,
      modifiedCols,
      code: 実列名.商品コード ? rowValues[列[実列名.商品コード] - 1] : '',
      sku: 実列名.SKU自動 ? rowValues[列[実列名.SKU自動] - 1] : '',
      status: 実列名.コードステータス ? rowValues[列[実列名.コードステータス] - 1] : ''
    });
  }

  if (works) {
    const 巻数 = 台湾書籍系_数字2桁_(値.単巻数);
    const セット終了 = 台湾書籍系_数字2桁_(値.セット終了);
    const 最新巻候補 = セット終了 ? parseInt(セット終了, 10) : (巻数 ? parseInt(巻数, 10) : null);

    // 【高速化】以前は フィールド更新_×3 + 予約巻数更新_ で Worksを4回フル読込していたのを
    // 1回の読込にまとめる（更新条件は従来と同一）。
    台湾書籍系_Worksまとめて更新_(設定, works.作品ID, {
      日本語タイトル: 台湾書籍系_正規化文字列_(値.日本語タイトル) ? 値.日本語タイトル : '',
      作者: 台湾書籍系_正規化文字列_(値.作者) ? 値.作者 : '',
      原題タイトル: 台湾書籍系_正規化文字列_(値.原題タイトル) ? 値.原題タイトル : '',
      最新巻: 最新巻候補,
    });
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

  // バッチ先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

  if (登録状況だけ編集) {
    if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
      台湾まんが_ONEDIT_LOG_('newOnEdit:status_skip', {
        sheet: sh.getName(),
        startRow: 開始行,
        rows: 行数,
        editedColumn: 編集開始列
      });
    }
    return;
  }

  let 発番発行TRUE行セット = null;
  if (発番発行だけ編集) {
    const values = e.range.getValues();
    const trueRows = [];
    for (let i = 0; i < values.length; i++) {
      const v = values[i] && values[i][0];
      if (v === true || String(v || '').toUpperCase() === 'TRUE') {
        trueRows.push(開始行 + i);
      }
    }

    if (!trueRows.length) {
      if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
        台湾まんが_ONEDIT_LOG_('newOnEdit:issue_checkbox_skip', {
          sheet: sh.getName(),
          startRow: 開始行,
          rows: 行数,
          editedColumn: 編集開始列,
          reason: 'no TRUE values'
        });
      }
      return;
    }

    発番発行TRUE行セット = {};
    trueRows.forEach(row => { 発番発行TRUE行セット[row] = true; });

    if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
      台湾まんが_ONEDIT_LOG_('newOnEdit:issue_checkbox_rows', {
        sheet: sh.getName(),
        rows: trueRows
      });
    }
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

  // バッチ先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

/* ============================================================
 * 取りこぼし補完（大量・高速編集対策）
 *  onEdit がロック取得失敗 / 行エラー / タイムアウトで補完できなかった行を
 *  保留キュー（DocumentProperties）に積み、約40秒後に1回動く時間トリガーで確実に補完する。
 *  補完(1行補完_共通_)は冪等なので、二重実行されても結果は変わらない。
 * ============================================================ */
var _台湾書籍系_保留KEY_ = '台湾書籍系_保留範囲v1';
var _台湾書籍系_保留トリガー名_ = '台湾書籍系_保留補完を処理_';

function 台湾書籍系_保留範囲をマージ_(list) {
  const bySheet = {};
  (list || []).forEach(function(r) {
    if (!r || !r.sheet) return;
    const s = Math.max(2, parseInt(r.start, 10) || 0);
    const e = parseInt(r.end, 10) || 0;
    if (e < s) return;
    (bySheet[r.sheet] = bySheet[r.sheet] || []).push([s, e]);
  });
  const out = [];
  Object.keys(bySheet).forEach(function(sheet) {
    const ranges = bySheet[sheet].sort(function(a, b) { return a[0] - b[0]; });
    let cur = null;
    ranges.forEach(function(pair) {
      const s = pair[0], e = pair[1];
      if (!cur) { cur = [s, e]; return; }
      if (s <= cur[1] + 1) { cur[1] = Math.max(cur[1], e); }
      else { out.push({ sheet: sheet, start: cur[0], end: cur[1] }); cur = [s, e]; }
    });
    if (cur) out.push({ sheet: sheet, start: cur[0], end: cur[1] });
  });
  return out;
}

function 台湾書籍系_保留に追加_(sheetName, startRow, numRows) {
  try {
    if (!sheetName || !numRows || numRows <= 0) return;
    const start = Math.max(2, startRow);
    const end = startRow + numRows - 1;
    if (end < start) return;
    const props = PropertiesService.getDocumentProperties();
    let list = [];
    try { list = JSON.parse(props.getProperty(_台湾書籍系_保留KEY_) || '[]'); } catch (_) { list = []; }
    if (!Array.isArray(list)) list = [];
    list.push({ sheet: sheetName, start: start, end: end });
    list = 台湾書籍系_保留範囲をマージ_(list);
    let json = JSON.stringify(list);
    if (json.length > 9000) json = JSON.stringify(list.slice(-50)); // プロパティ上限対策
    props.setProperty(_台湾書籍系_保留KEY_, json);
  } catch (err) {
    デバッグログ出力_('保留に追加_エラー', { message: String(err) });
  }
}

function 台湾書籍系_保留補完トリガーを確保_() {
  try {
    const exists = ScriptApp.getProjectTriggers()
      .some(function(t) { return t.getHandlerFunction() === _台湾書籍系_保留トリガー名_; });
    if (exists) return;
    ScriptApp.newTrigger(_台湾書籍系_保留トリガー名_).timeBased().after(40 * 1000).create();
  } catch (err) {
    デバッグログ出力_('保留トリガー確保_エラー', { message: String(err) });
  }
}

function 台湾書籍系_保留設定マップ_() {
  const m = {};
  if (typeof 設定_台湾まんが !== 'undefined' && 設定_台湾まんが) m[設定_台湾まんが.マスターシート名] = 設定_台湾まんが;
  if (typeof 設定_台湾書籍その他 !== 'undefined' && 設定_台湾書籍その他) m[設定_台湾書籍その他.マスターシート名] = 設定_台湾書籍その他;
  return m;
}

function 台湾書籍系_保留補完を処理_() {
  // 一回限り(after)のトリガーなので、まず自分を掃除する
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === _台湾書籍系_保留トリガー名_; })
      .forEach(function(t) { try { ScriptApp.deleteTrigger(t); } catch (_) {} });
  } catch (_) {}

  const props = PropertiesService.getDocumentProperties();
  let list = [];
  try { list = JSON.parse(props.getProperty(_台湾書籍系_保留KEY_) || '[]'); } catch (_) { list = []; }
  if (!Array.isArray(list) || !list.length) {
    try { props.deleteProperty(_台湾書籍系_保留KEY_); } catch (_) {}
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) {
    // まだ他処理がロック中。改めてトリガーを張って後で再試行（キューは消さない）。
    台湾書籍系_保留補完トリガーを確保_();
    return;
  }

  try {
    台湾書籍系_使用済みIDキャッシュ無効化_();
    const ss = SpreadsheetApp.getActive();
    const 設定マップ = 台湾書籍系_保留設定マップ_();

    list.forEach(function(r) {
      const 設定 = 設定マップ[r.sheet];
      if (!設定) return;
      const sh = ss.getSheetByName(r.sheet);
      if (!sh) return;
      const 列 = 台湾書籍系_列マップを取得_(sh);
      const 登録状況列 = 列['登録状況'];
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      const start = Math.max(2, r.start);
      const end = Math.min(r.end, lastRow);
      if (end < start) return;
      const allRowValues = sh.getRange(start, 1, end - start + 1, lastCol).getValues();
      for (let row = start; row <= end; row++) {
        const rowValues = allRowValues[row - start];
        if (!rowValues) continue;
        const 状況 = 登録状況列 ? String(rowValues[登録状況列 - 1] || '').trim() : '';
        if (状況.startsWith('登録済')) continue; // 登録済みは触らない
        try {
          台湾書籍系_1行補完_共通_(sh, row, 設定, {
            Works新規作成: true,
            skipFlush: true,
            列マップ: 列,
            lastCol: lastCol,
            rowValues: rowValues
          });
        } catch (rowErr) {
          デバッグログ出力_('保留補完_行エラー', { sheet: r.sheet, row: row, message: String(rowErr) });
        }
      }
    });
    try { props.deleteProperty(_台湾書籍系_保留KEY_); } catch (_) {}
    SpreadsheetApp.flush();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function 台湾書籍系_新onEdit_共通_(e, 設定) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== 設定.マスターシート名) return;

  // この実行バッチの先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

  if (登録状況だけ編集) {
    if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
      台湾まんが_ONEDIT_LOG_('newOnEdit:status_skip', {
        sheet: sh.getName(),
        startRow: 開始行,
        rows: 行数,
        editedColumn: 編集開始列
      });
    }
    return;
  }

  let 発番発行TRUE行セット = null;
  if (発番発行だけ編集) {
    const values = e.range.getValues();
    const trueRows = [];
    for (let i = 0; i < values.length; i++) {
      const v = values[i] && values[i][0];
      if (v === true || String(v || '').toUpperCase() === 'TRUE') {
        trueRows.push(開始行 + i);
      }
    }

    if (!trueRows.length) {
      if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
        台湾まんが_ONEDIT_LOG_('newOnEdit:issue_checkbox_skip', {
          sheet: sh.getName(),
          startRow: 開始行,
          rows: 行数,
          editedColumn: 編集開始列,
          reason: 'no TRUE values'
        });
      }
      return;
    }

    発番発行TRUE行セット = {};
    trueRows.forEach(row => { 発番発行TRUE行セット[row] = true; });

    if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
      台湾まんが_ONEDIT_LOG_('newOnEdit:issue_checkbox_rows', {
        sheet: sh.getName(),
        rows: trueRows
      });
    }
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

  const 監視列編集あり = 監視列番号.some(col =>
    col >= 編集開始列 && col <= 編集終了列
  );
  // 既存値の上から貼り付けた場合、編集列が監視列だけに限定できないことがある。
  // A/Bの操作列以外なら未登録行を再計算し、D列の巻数反映漏れを拾う。
  const データ列編集あり = 編集開始列 >= 3 || 編集終了列 >= 3;
  const 対象列あり = !!発番発行TRUE行セット || 監視列編集あり || データ列編集あり;

  デバッグログ出力_('書籍系_新onEdit_共通_到達', {
    sheet: sh.getName(),
    開始行,
    行数,
    編集開始列,
    編集終了列,
    対象列あり,
    Works新規作成: true
  });
  if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
    台湾まんが_ONEDIT_LOG_('newOnEdit:watch_check', {
      sheet: sh.getName(),
      startRow: 開始行,
      rows: 行数,
      editedStartColumn: 編集開始列,
      editedEndColumn: 編集終了列,
      watchedColumns: 監視列番号,
      matched: 対象列あり,
      watchedMatched: 監視列編集あり,
      dataColumnFallback: データ列編集あり,
      codeColumn: 列['親コード'] || 0,
      skuColumn: 列['SKU（自動）'] || 列['SKU(自動)'] || 0,
      volumeColumn: 列['単巻数'] || 0,
      setStartColumn: 列['セット巻数開始番号'] || 0,
      setEndColumn: 列['セット巻数終了番号'] || 0
    });
  }

  if (!対象列あり) return;

  // 【取りこぼし防止・大量入力対策】大きな貼り付け/編集（しきい値以上の行数）は、
  // 処理前に保留キューへ積んでおく。万一インライン処理が6分制限でタイムアウトしても、
  // 後続の時間トリガーが未補完行を拾える。補完は冪等なので、正常完了後に再検証されても無害。
  if (!発番発行TRUE行セット && 行数 >= 10) {
    台湾書籍系_保留に追加_(sh.getName(), 開始行, 行数);
    台湾書籍系_保留補完トリガーを確保_();
  }

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    // 【取りこぼし防止】黙って捨てず、対象行を保留キューに積んで後で時間トリガーで補完する。
    // 大量・高速編集でロックが取れず処理スキップ → 「反映されない/消える」の主因だった。
    台湾書籍系_保留に追加_(sh.getName(), 開始行, 行数);
    台湾書籍系_保留補完トリガーを確保_();
    return;
  }

  // ループ前に数式を確定し、必要な情報を一括取得
  SpreadsheetApp.flush();

  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(開始行, 1, 行数, lastCol).getValues();

  let エラー行あり = false;
  try {
    for (let r = 開始行; r < 開始行 + 行数; r++) {
      if (r < 2) continue;
      if (発番発行TRUE行セット && !発番発行TRUE行セット[r]) continue;

      const rowValues = allRowValues[r - 開始行];
      if (!rowValues) continue;

      const 登録状況値 = 登録状況列
        ? String(rowValues[登録状況列 - 1] || '').trim()
        : '';

      // 登録済み行は、通常のonEditでは自動補完・再採番しない
      if (登録状況値.startsWith('登録済')) {
        if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
          台湾まんが_ONEDIT_LOG_('newOnEdit:registered_skip', {
            sheet: sh.getName(),
            row: r,
            status: 登録状況値
          });
        }
        continue;
      }

      // 1行の失敗で残り行を巻き添えにしない（部分的な未反映を防ぐ）
      try {
        台湾書籍系_1行補完_共通_(sh, r, 設定, {
          Works新規作成: true,
          skipFlush: true,
          列マップ: 列,
          lastCol: lastCol,
          rowValues: rowValues
        });
      } catch (rowErr) {
        エラー行あり = true;
        デバッグログ出力_('新onEdit_行補完エラー', { row: r, message: String(rowErr) });
      }
    }
  } finally {
    lock.releaseLock();
  }

  // 一部の行でエラーが出ていたら、保留に積んで後で再補完（取りこぼし防止）。
  if (エラー行あり) {
    台湾書籍系_保留に追加_(sh.getName(), 開始行, 行数);
    台湾書籍系_保留補完トリガーを確保_();
  }
}

function 台湾書籍系_新取込後補完_共通_(sheetName, startRow, numRows, cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  if (!sh || !numRows || numRows <= 0) return;

  // バッチ先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

  // バッチ先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

// 1実行内で Works作品ID列の '@' 書式を二重設定しないためのガード
var _台湾書籍系_ID列書式済み_ = (typeof _台湾書籍系_ID列書式済み_ !== 'undefined') ? _台湾書籍系_ID列書式済み_ : {};
function 台湾書籍系_Works作品ID列を文字列固定_(sh, シート名) {
  if (!sh) return;
  if (_台湾書籍系_ID列書式済み_[シート名]) return;
  if (sh.getMaxRows() >= 2) {
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  _台湾書籍系_ID列書式済み_[シート名] = true;
}

function 台湾書籍系_Worksを取得または作成_(設定, 値) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定.作品シート名);

  if (!sh) {
    sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品列数).setValues([設定.作品ヘッダー]);
  }

  // Worksの作品ID列は常に文字列固定（これをしないと 0087 が 87 に落ちる）。
  // ★高速化: 以前は呼ばれるたびにシート全体へ書式設定していたため、
  //   一括更新では行数分くり返され非常に遅かった。1実行につき1回だけ実行する。
  台湾書籍系_Works作品ID列を文字列固定_(sh, 設定.作品シート名);

  const 日本語タイトル = 台湾書籍系_正規化文字列_(値.日本語タイトル);
  const 作者 = 台湾書籍系_正規化文字列_(値.作者);
  const 原題タイトル = 台湾書籍系_正規化文字列_(値.原題タイトル || 値.原題商品タイトル);
  const 媒体コード = 台湾書籍系_媒体コードを生成_(値, 設定 && 設定.マスターシート名);
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

const found = 台湾書籍系_Worksから同一作品情報を取得_(設定, Object.assign({}, 値, {
  原題タイトル,
  日本語タイトル,
  作者
})) || 台湾書籍系_Worksから取得_(
  設定,
  原題タイトル,
  日本語タイトル,
  作者,
  媒体コード
);



if (found && 台湾書籍系_作品ID4桁を取得_(found.作品ID)) {
  return {
    作品ID: 台湾書籍系_作品ID4桁を取得_(found.作品ID),
    日本語タイトル: found.日本語タイトル || 日本語タイトル,
    作者: found.作者 || 作者,
    原題タイトル: found.原題タイトル || 原題タイトル
  };
}

const 作品名 = 日本語タイトル || 原題タイトル;
if (!作品名) {
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
    : `${作品名}||${作者}`;

  const worksKey = 原題タイトル
    ? (媒体コード ? `原題||${媒体コード}||${worksKeyBase}` : `原題||${worksKeyBase}`)
    : (媒体コード ? `${媒体コード}||${worksKeyBase}` : worksKeyBase);

  const writeRow = Math.max(sh.getLastRow() + 1, 2);

  // ★ここに入れる
  デバッグログ出力_('Works新規作成_実行', {
    新作品ID,
    worksKey,
    日本語タイトル: 作品名,
    作者,
    原題タイトル,
    writeRow
  });

  // 作品ID列は、書き込み直前にも文字列固定
  sh.getRange(writeRow, 2).setNumberFormat('@');

  sh.getRange(writeRow, 1, 1, 設定.作品列数).setValues([[
    worksKey,
    新作品ID,
    作品名,
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
    日本語タイトル: 作品名,
    作者,
    原題タイトル
  };
}

function 台湾書籍系_WorksKey媒体付与_削除なし() {
  const ui = SpreadsheetApp.getUi();
  if (
    ui.alert(
      '確認',
      'WorksKeyにCM/NVを付与します。\n\n' +
      '・台湾まんがで使われている作品IDは CM\n' +
      '・台湾書籍その他でカテゴリが小説/ノベルの作品IDは NV\n' +
      '・Works行の削除、統合、作品ID変更は一切しません\n' +
      '・判定できない行はスキップします\n\n' +
      '続行しますか？',
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const result = 台湾書籍系_WorksKey媒体付与_削除なし_実行_();
    ui.alert(
      '完了',
      'WorksKeyにCM/NVを付与しました（削除なし）。\n\n' +
      `更新: ${result.更新数}件\n` +
      `既に媒体付き: ${result.既存媒体付き数}件\n` +
      `使用媒体が不明でスキップ: ${result.媒体不明数}件\n` +
      `CM/NV以外でスキップ: ${result.CMNV以外数}件\n` +
      `同一IDでCM/NV混在のためスキップ: ${result.媒体混在数}件\n` +
      `既存媒体と使用媒体が不一致でスキップ: ${result.媒体不一致数}件\n\n` +
      '詳細はApps Scriptの実行ログに出しています。',
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍系_WorksKey媒体付与_削除なし_実行_() {
  const ss = SpreadsheetApp.getActive();
  const 作品シート名 =
    (typeof 設定_台湾まんが !== 'undefined' && 設定_台湾まんが && 設定_台湾まんが.作品シート名) ||
    (typeof 設定_台湾書籍その他 !== 'undefined' && 設定_台湾書籍その他 && 設定_台湾書籍その他.作品シート名) ||
    'Works（書籍専用）';

  const worksSh = ss.getSheetByName(作品シート名);
  if (!worksSh || worksSh.getLastRow() < 2) {
    throw new Error(`${作品シート名} にデータがありません`);
  }

  const idToMediaSet = {};
  const usageStats = {
    台湾まんが: 0,
    台湾書籍その他: 0
  };

  const collect = (設定, fallbackSheetName) => {
    if (!設定) return;
    const シート名 = 設定.マスターシート名 || fallbackSheetName;
    const sh = ss.getSheetByName(シート名);
    if (!sh || sh.getLastRow() < 2) return;

    const 列 = 台湾書籍系_列マップを取得_(sh);
    const cn = 設定.列名 || {};
    const 実列名 = {
      商品コード: 台湾書籍系_実列名を取得_(列, [cn.商品コード, '親コード', '商品コード', '商品コード(SKU)', '商品コード（SKU）']),
      作品ID: 台湾書籍系_実列名を取得_(列, [cn.作品ID, '作品ID(W)(自動)', '作品ID(W)（自動）', '作品(W)（自動）']),
      SKU自動: 台湾書籍系_実列名を取得_(列, [cn.SKU自動, 'SKU(自動)', 'SKU（自動）', '商品コード(SKU)', '商品コード（SKU）']),
      カテゴリ: 台湾書籍系_実列名を取得_(列, [cn.カテゴリ, 'カテゴリ'])
    };

    const lastCol = sh.getLastColumn();
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    values.forEach(row => {
      const get = (列名) => (列名 && 列[列名]) ? row[列[列名] - 1] : '';
      const 値 = {
        作品ID: get(実列名.作品ID),
        商品コード: get(実列名.商品コード),
        親コード: get(実列名.商品コード),
        SKU自動: get(実列名.SKU自動),
        カテゴリ: get(実列名.カテゴリ)
      };

      const id = 台湾書籍系_行から既存作品IDを取得_(値);
      if (!id) return;

      const media = 台湾書籍系_媒体コードを生成_(値, シート名);
      if (!media) return;

      idToMediaSet[id] = idToMediaSet[id] || {};
      idToMediaSet[id][media] = true;
      if (usageStats[シート名] != null) usageStats[シート名]++;
    });
  };

  collect(typeof 設定_台湾まんが !== 'undefined' ? 設定_台湾まんが : null, '台湾まんが');
  collect(typeof 設定_台湾書籍その他 !== 'undefined' ? 設定_台湾書籍その他 : null, '台湾書籍その他');

  const worksCols = 台湾書籍系_列マップを取得_(worksSh);
  const colWorksKey = worksCols['WorksKey'];
  const col作品ID = worksCols['作品ID'];
  const col日本語タイトル = worksCols['日本語タイトル'];
  const col作者 = worksCols['作者'];
  const col原題タイトル = worksCols['原題タイトル'];

  if (!colWorksKey || !col作品ID) {
    throw new Error('Worksに WorksKey / 作品ID 列が見つかりません');
  }

  const lastRow = worksSh.getLastRow();
  const lastCol = worksSh.getLastColumn();
  const rows = worksSh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const updates = [];
  const skipped = {
    既存媒体付き数: 0,
    媒体不明数: 0,
    CMNV以外数: 0,
    媒体混在数: 0,
    媒体不一致数: 0
  };
  const detail = {
    媒体混在ID: [],
    媒体不一致行: [],
    媒体不明ID: []
  };

  rows.forEach((row, i) => {
    const sheetRow = i + 2;
    const id = 台湾書籍系_作品ID4桁を取得_(row[col作品ID - 1]);
    if (!id) {
      skipped.媒体不明数++;
      return;
    }

    const mediaSet = idToMediaSet[id] || {};
    const medias = Object.keys(mediaSet).filter(Boolean);
    const currentKey = 台湾書籍系_正規化文字列_(row[colWorksKey - 1]);
    const currentMedia = 台湾書籍系_WorksKey媒体コード_(currentKey);

    if (!medias.length) {
      skipped.媒体不明数++;
      if (detail.媒体不明ID.length < 20) detail.媒体不明ID.push(id);
      return;
    }

    if (medias.length > 1) {
      skipped.媒体混在数++;
      if (detail.媒体混在ID.length < 20) detail.媒体混在ID.push(`${id}: ${medias.join('/')}`);
      return;
    }

    const media = medias[0];
    if (media !== 'CM' && media !== 'NV') {
      skipped.CMNV以外数++;
      return;
    }

    if (currentMedia) {
      if (currentMedia === media) {
        skipped.既存媒体付き数++;
      } else {
        skipped.媒体不一致数++;
        if (detail.媒体不一致行.length < 20) {
          detail.媒体不一致行.push(`${sheetRow}行目 ID:${id} WorksKey:${currentMedia} 使用:${media}`);
        }
      }
      return;
    }

    const newKey = 台湾書籍系_WorksKey媒体付きへ変換_(
      currentKey,
      media,
      col原題タイトル ? row[col原題タイトル - 1] : '',
      col日本語タイトル ? row[col日本語タイトル - 1] : '',
      col作者 ? row[col作者 - 1] : ''
    );

    if (!newKey || newKey === currentKey) return;
    updates.push({ row: sheetRow, key: newKey });
  });

  updates.forEach(item => {
    worksSh.getRange(item.row, colWorksKey).setValue(item.key);
  });

  Logger.log('WorksKey媒体付与_削除なし usageStats=' + JSON.stringify(usageStats));
  Logger.log('WorksKey媒体付与_削除なし updates=' + JSON.stringify(updates.slice(0, 50)));
  Logger.log('WorksKey媒体付与_削除なし skipped=' + JSON.stringify(skipped));
  Logger.log('WorksKey媒体付与_削除なし detail=' + JSON.stringify(detail));

  return {
    更新数: updates.length,
    既存媒体付き数: skipped.既存媒体付き数,
    媒体不明数: skipped.媒体不明数,
    CMNV以外数: skipped.CMNV以外数,
    媒体混在数: skipped.媒体混在数,
    媒体不一致数: skipped.媒体不一致数
  };
}

function 台湾書籍系_WorksKey媒体付きへ変換_(currentKey, media, 原題タイトル, 日本語タイトル, 作者) {
  const key = 台湾書籍系_正規化文字列_(currentKey);
  const m = 台湾書籍系_正規化文字列_(media).toUpperCase();
  if (!m) return key;
  if (台湾書籍系_WorksKey媒体コード_(key)) return key;

  const original = 台湾書籍系_正規化文字列_(原題タイトル);
  if (original) {
    const base = 台湾書籍系_作品比較キー_(original) || original;
    return `原題||${m}||${base}`;
  }

  if (key.indexOf('原題||') === 0) {
    const base = key.split('||').slice(1).join('||');
    if (base) return `原題||${m}||${base}`;
  }

  const fallbackBase = key || [日本語タイトル, 作者]
    .map(v => 台湾書籍系_正規化文字列_(v))
    .filter(Boolean)
    .join('||');

  return fallbackBase ? `${m}||${fallbackBase}` : '';
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

/**
 * 【onEdit/一括処理 高速化】使用済み作品IDセットの実行内キャッシュ。
 * 使用済み作品IDセット_() は全商品シート＋Worksを走査するため非常に重く、
 * 大量の新規採番で行ごとに走査すると致命的に遅い。1実行内では1回だけ構築して使い回す。
 * ※インスタンス再利用による実行跨ぎの陳腐化を避けるため、各実行の入口で必ず無効化する。
 */
var _台湾書籍系_使用済みIDキャッシュ_ = (typeof _台湾書籍系_使用済みIDキャッシュ_ !== 'undefined') ? _台湾書籍系_使用済みIDキャッシュ_ : null;

function 台湾書籍系_使用済みIDキャッシュ無効化_() {
  _台湾書籍系_使用済みIDキャッシュ_ = null;
}

function 台湾書籍系_使用済みIDキャッシュ取得_() {
  if (_台湾書籍系_使用済みIDキャッシュ_ instanceof Set) return _台湾書籍系_使用済みIDキャッシュ_;
  _台湾書籍系_使用済みIDキャッシュ_ = 台湾書籍系_使用済み作品IDセット_();
  return _台湾書籍系_使用済みIDキャッシュ_;
}

function 台湾書籍系_次の未使用作品ID_() {
  // 実行内キャッシュを使用（全シート走査を実行ごとに1回へ削減）
  const used = 台湾書籍系_使用済みIDキャッシュ取得_();

  // 0001から順に、本当に空いている最小IDを使う
  for (let n = 1; n < 10000; n++) {
    const id = String(n).padStart(4, '0');

    if (!used.has(id)) {
      // 同一実行内で次の採番に同じIDを返さないよう、即座に予約（衝突防止）
      used.add(id);
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

  // バッチ先頭で使用済みIDキャッシュを無効化（実行跨ぎの陳腐化＝ID衝突を防ぐ）。
  台湾書籍系_使用済みIDキャッシュ無効化_();

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

/**
 * Works の複数フィールド（日本語タイトル/作者/原題タイトル/最新巻）を
 * 1回の読込でまとめて更新する。【onEdit高速化】
 * 以前は フィールド更新_×3 + 予約巻数更新_ で Worksシートを4回フル読込していたが、
 * 値が大きいと1行ごとに数万セルを4回読むため非常に遅かった。ここでは1回だけ読込み、
 * 該当行に対して必要なセルだけを書き込む。ロジック（更新条件）は従来と完全に同一。
 * updates: { 日本語タイトル, 作者, 原題タイトル, 最新巻 }（不要なものは空文字 / null）
 */
function 台湾書籍系_Worksまとめて更新_(設定, 作品ID, updates) {
  const id = 台湾書籍系_作品ID4桁を取得_(作品ID);
  if (!id || !updates) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(v => 台湾書籍系_正規化文字列_(v));
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  if (col['作品ID'] == null) return;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();

  for (let i = 0; i < data.length; i++) {
    if (台湾書籍系_作品ID4桁を取得_(data[i][col['作品ID']]) !== id) continue;
    const rowNo = i + 2;

    // テキスト系: 従来の フィールド更新_ と同じく「現在値と異なる時のみ」上書き
    [
      ['日本語タイトル', updates.日本語タイトル],
      ['作者', updates.作者],
      ['原題タイトル', updates.原題タイトル],
    ].forEach(([フィールド名, 新値]) => {
      if (新値 == null || 新値 === '') return;
      if (col[フィールド名] == null) return;
      const 現在値 = 台湾書籍系_正規化文字列_(data[i][col[フィールド名]]);
      const 新値正規化 = 台湾書籍系_正規化文字列_(新値);
      if (現在値 !== 新値正規化) {
        sh.getRange(rowNo, col[フィールド名] + 1).setValue(新値);
      }
    });

    // 最新巻(予約込み): 従来の 予約巻数更新_ と同じく「巻数 > 既存」の時のみ
    if (updates.最新巻 != null && col['最新巻(予約込み)'] != null) {
      const 既存 = parseInt(String(data[i][col['最新巻(予約込み)']] || '0'), 10) || 0;
      if (updates.最新巻 > 既存) {
        sh.getRange(rowNo, col['最新巻(予約込み)'] + 1).setValue(updates.最新巻);
        if (col['予約更新日時'] != null) {
          sh.getRange(rowNo, col['予約更新日時'] + 1).setValue(new Date());
        }
      }
    }
    return;
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
