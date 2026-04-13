/* ============================================================
 * 台湾書籍系 段階生成 共通
 * ============================================================ */

function 台湾書籍系_列マップを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  return map;
}

function 台湾書籍系_実列名を取得_(列, 候補配列) {
  for (const name of 候補配列) {
    if (name && 列[name]) return name;
  }
  return '';
}

function 台湾書籍系_Worksから取得_(設定, 原題タイトル) {
  const key = String(原題タイトル || '').trim();
  if (!key) return null;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.作品シート名);
  if (!sh || sh.getLastRow() < 2) return null;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (const row of data) {
    const raw = String(row[col['原題タイトル']] || '').trim();
    if (raw === key) {
      return {
        作品ID:       col['作品ID'] != null ? row[col['作品ID']] : '',
        日本語タイトル: col['日本語タイトル'] != null ? row[col['日本語タイトル']] : '',
        作者:         col['作者'] != null ? row[col['作者']] : ''
      };
    }
  }
  return null;
}

function 台湾書籍系_言語表示を生成_(言語) {
  const v = String(言語 || '').trim();
  if (!v) return '';
  if (v === '台湾' || v.toUpperCase() === 'TW') return '台湾版';
  if (v === '中国' || v.toUpperCase() === 'CN') return '中国版';
  if (v === '韓国' || v.toUpperCase() === 'KR') return '韓国版';
  return `${v}版`;
}

function 台湾書籍系_カテゴリ表示を生成_(シート名, カテゴリ) {
  const v = String(カテゴリ || '').trim();
  if (v) return v;
  return シート名 === '台湾まんが' ? 'まんが' : '書籍';
}

function 台湾書籍系_巻表示を生成_(単巻数, セット開始, セット終了) {
  const 単 = String(単巻数 || '').trim();
  const 開 = String(セット開始 || '').trim();
  const 終 = String(セット終了 || '').trim();

  if (開 && 終) {
    if (開 === 終) return `${開}巻`;
    return `${開}〜${終}巻セット`;
  }
  if (単) return `${単}巻`;
  return '';
}

function 台湾書籍系_タイトルを生成_(シート名, 値) {
  const 言語表記 = 台湾書籍系_言語表示を生成_(値.言語);
  const カテゴリ = 台湾書籍系_カテゴリ表示を生成_(シート名, 値.カテゴリ);
  const ベースタイトル = String(値.日本語タイトル || 値.原題タイトル || '').trim();
  const 巻表示 = 台湾書籍系_巻表示を生成_(値.単巻数, 値.セット開始, 値.セット終了);
  const 形態 = String(値.形態 || '').trim();
  const 特典メモ = String(値.特典メモ || '').trim();

  if (!ベースタイトル) return '';

  const parts = [];
  if (言語表記) parts.push(言語表記);
  if (カテゴリ) parts.push(カテゴリ);

  let title = parts.join(' ');
  title += ` 『${ベースタイトル}』`;

  const tail = [];
  if (巻表示) tail.push(巻表示);
  if (形態 && 形態 !== '通常') tail.push(形態);
  if (特典メモ) tail.push(特典メモ);

  if (tail.length) title += ' ' + tail.join(' ');
  return title.replace(/\s+/g, ' ').trim();
}

function 台湾書籍系_重複チェックキーを生成_(値) {
  const サイト商品コード = String(値.サイト商品コード || '').trim();
  const isbn = String(値.ISBN || '').replace(/[^0-9Xx]/g, '');
  const 原題 = String(値.原題タイトル || '').trim();
  const 日本語 = String(値.日本語タイトル || '').trim();
  const 形態 = String(値.形態 || '').trim();
  const 巻表示 = 台湾書籍系_巻表示を生成_(値.単巻数, 値.セット開始, 値.セット終了);

  if (サイト商品コード) return `SITE||${サイト商品コード}`;
  if (isbn) return `ISBN||${isbn}`;

  const base = 原題 || 日本語;
  if (!base) return '';

  return [base, 巻表示, 形態].filter(Boolean).join('||');
}

function 台湾書籍系_粗利益率を計算_(売価, 原価) {
  const 売 = parseFloat(売価 || 0);
  const 原 = parseFloat(原価 || 0);
  if (!(売 > 0 && 原 > 0)) return null;

  let レート = 0;
  try {
    レート = _kyoutuu.為替レートを取得_('TWD');
  } catch (_) {
    レート = 0;
  }
  if (!(レート > 0)) return null;

  return Math.round(((売 - 原 * レート) / 売) * 1000) / 1000;
}

function 台湾書籍系_不足項目を返す_(シート名, 値) {
  const lacks = [];
  if (!String(値.言語 || '').trim()) lacks.push('言語');
  if (!String(値.日本語タイトル || 値.原題タイトル || '').trim()) lacks.push('タイトル');
  if (!String(値.形態 || '').trim()) lacks.push('形態');

  const カテゴリ = 台湾書籍系_カテゴリ表示を生成_(シート名, 値.カテゴリ);
  if (!カテゴリ) lacks.push('カテゴリ');

  return lacks;
}

function 台湾書籍系_関数を探す_(候補名配列) {
  for (const name of 候補名配列) {
    try {
      const fn = globalThis[name];
      if (typeof fn === 'function') return fn;
    } catch (_) {}
  }
  return null;
}

function 台湾書籍系_親コードSKUを試行生成_(sh, row, 設定, 列, rowValues, 値, 実列名) {
  let changed = false;

  const 親コード関数 = 台湾書籍系_関数を探す_([
    `${sh.getName()}_親コードを生成_`,
    '台湾書籍系_親コードを生成_共通_',
    '書籍系_親コードを生成_共通_'
  ]);

  const SKU関数 = 台湾書籍系_関数を探す_([
    `${sh.getName()}_SKUを生成_`,
    '台湾書籍系_SKUを生成_共通_',
    '書籍系_SKUを生成_共通_'
  ]);

  const 商品コード列名 = 実列名.商品コード;
  const SKU列名 = 実列名.SKU自動;
  const ステータス列名 = 実列名.コードステータス;

  const 現在親コード = 商品コード列名 ? String(rowValues[列[商品コード列名] - 1] || '').trim() : '';
  const 現在SKU = SKU列名 ? String(rowValues[列[SKU列名] - 1] || '').trim() : '';

  if (親コード関数 && 商品コード列名 && !現在親コード) {
    try {
      const code = 親コード関数({
        sheet: sh,
        row,
        設定,
        列,
        値
      });
      if (code) {
        rowValues[列[商品コード列名] - 1] = code;
        changed = true;
      }
    } catch (err) {
      if (ステータス列名) rowValues[列[ステータス列名] - 1] = `親コード生成エラー: ${err}`;
      changed = true;
    }
  }

  if (SKU関数 && SKU列名 && !現在SKU) {
    try {
      const sku = SKU関数({
        sheet: sh,
        row,
        設定,
        列,
        値,
        親コード: 商品コード列名 ? rowValues[列[商品コード列名] - 1] : ''
      });
      if (sku) {
        rowValues[列[SKU列名] - 1] = sku;
        changed = true;
      }
    } catch (err) {
      if (ステータス列名) rowValues[列[ステータス列名] - 1] = `SKU生成エラー: ${err}`;
      changed = true;
    }
  }

  return changed;
}

function 台湾書籍系_1行補完_共通_(sh, row, 設定) {
  if (!sh || row < 2) return;

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  const 実列名 = {
    商品コード: 台湾書籍系_実列名を取得_(列, [設定.列名.商品コード, '親コード', '商品コード(SKU)']),
    タイトル: 台湾書籍系_実列名を取得_(列, [設定.列名.タイトル, 'タイトル']),
    作者: 台湾書籍系_実列名を取得_(列, [設定.列名.作者, '作者']),
    日本語タイトル: 台湾書籍系_実列名を取得_(列, [設定.列名.日本語タイトル, '日本語タイトル']),
    原題: 台湾書籍系_実列名を取得_(列, [設定.列名.原題, '原題タイトル']),
    形態: 台湾書籍系_実列名を取得_(列, [設定.列名.形態, '形態(通常/初回限定/特装)']),
    言語: 台湾書籍系_実列名を取得_(列, [設定.列名.言語, '言語']),
    カテゴリ: 台湾書籍系_実列名を取得_(列, [設定.列名.カテゴリ, 'カテゴリ']),
    単巻数: 台湾書籍系_実列名を取得_(列, [設定.列名.単巻数, '単巻数']),
    セット開始: 台湾書籍系_実列名を取得_(列, [設定.列名.セット開始, 'セット巻数開始番号']),
    セット終了: 台湾書籍系_実列名を取得_(列, [設定.列名.セット終了, 'セット巻数終了番号']),
    特典メモ: 台湾書籍系_実列名を取得_(列, [設定.列名.特典メモ, '特典メモ']),
    ISBN: 台湾書籍系_実列名を取得_(列, [設定.列名.ISBN, 'ISBN']),
    作品ID: 台湾書籍系_実列名を取得_(列, [設定.列名.作品ID, '作品ID(W)(自動)']),
    SKU自動: 台湾書籍系_実列名を取得_(列, [設定.列名.SKU自動, 'SKU(自動)']),
    コードステータス: 台湾書籍系_実列名を取得_(列, [設定.列名.コードステータス, '商品コードステータス']),
    発行チェック: 台湾書籍系_実列名を取得_(列, [設定.列名.発行チェック, '発番発行']),
    登録状況: 台湾書籍系_実列名を取得_(列, [設定.列名.登録状況, '登録状況']),
    サイト商品コード: 台湾書籍系_実列名を取得_(列, [設定.列名.博客來商品コード, 'サイト商品コード', '博客來商品コード']),
    売価: 台湾書籍系_実列名を取得_(列, ['売価']),
    原価: 台湾書籍系_実列名を取得_(列, ['原価']),
    粗利益率: 台湾書籍系_実列名を取得_(列, ['粗利益率']),
    重複チェックキー: 台湾書籍系_実列名を取得_(列, ['重複チェックキー'])
  };

  const g = (列名) => 列名 ? rowValues[列[列名] - 1] : '';
  const s = (列名, 値) => {
    if (!列名) return false;
    const idx = 列[列名] - 1;
    if (rowValues[idx] === 値) return false;
    rowValues[idx] = 値;
    return true;
  };

  let changed = false;

  let 値 = {
    作者: g(実列名.作者),
    日本語タイトル: g(実列名.日本語タイトル),
    原題タイトル: g(実列名.原題),
    形態: g(実列名.形態),
    言語: g(実列名.言語),
    カテゴリ: g(実列名.カテゴリ),
    単巻数: g(実列名.単巻数),
    セット開始: g(実列名.セット開始),
    セット終了: g(実列名.セット終了),
    特典メモ: g(実列名.特典メモ),
    ISBN: g(実列名.ISBN),
    サイト商品コード: g(実列名.サイト商品コード)
  };

  // 1) Works から補完
  const works = 台湾書籍系_Worksから取得_(設定, 値.原題タイトル);
  if (works) {
    if (!String(値.日本語タイトル || '').trim() && String(works.日本語タイトル || '').trim()) {
      changed = s(実列名.日本語タイトル, works.日本語タイトル) || changed;
      値.日本語タイトル = works.日本語タイトル;
    }
    if (!String(値.作者 || '').trim() && String(works.作者 || '').trim()) {
      changed = s(実列名.作者, works.作者) || changed;
      値.作者 = works.作者;
    }
    if (!String(g(実列名.作品ID) || '').trim() && String(works.作品ID || '').trim()) {
      changed = s(実列名.作品ID, works.作品ID) || changed;
    }
  }

  // 2) タイトル
  const タイトル = 台湾書籍系_タイトルを生成_(sh.getName(), 値);
  if (タイトル) {
    changed = s(実列名.タイトル, タイトル) || changed;
  }

  // 3) 重複チェックキー
  const 重複キー = 台湾書籍系_重複チェックキーを生成_(値);
  if (重複キー) {
    changed = s(実列名.重複チェックキー, 重複キー) || changed;
  }

  // 4) 粗利益率
  const 粗利益率 = 台湾書籍系_粗利益率を計算_(g(実列名.売価), g(実列名.原価));
  if (粗利益率 != null && 実列名.粗利益率) {
    changed = s(実列名.粗利益率, 粗利益率) || changed;
  }

  // 5) 商品コードステータス / 発番
  const 不足 = 台湾書籍系_不足項目を返す_(sh.getName(), 値);
  const 発番チェック = g(実列名.発行チェック) === true;

  if (不足.length > 0) {
    changed = s(実列名.コードステータス, '情報不足: ' + 不足.join(',')) || changed;
  } else {
    if (発番チェック) {
      const codeChanged = 台湾書籍系_親コードSKUを試行生成_(sh, row, 設定, 列, rowValues, 値, 実列名);
      changed = changed || codeChanged;

      const 親コード = 実列名.商品コード ? String(rowValues[列[実列名.商品コード] - 1] || '').trim() : '';
      const SKU = 実列名.SKU自動 ? String(rowValues[列[実列名.SKU自動] - 1] || '').trim() : '';

      if (親コード || SKU) {
        changed = s(実列名.コードステータス, '生成済み') || changed;
      } else {
        changed = s(実列名.コードステータス, '発番待ち') || changed;
      }
    } else {
      const 親コード = 実列名.商品コード ? String(g(実列名.商品コード) || '').trim() : '';
      const SKU = 実列名.SKU自動 ? String(g(実列名.SKU自動) || '').trim() : '';
      if (親コード || SKU) {
        changed = s(実列名.コードステータス, '生成済み') || changed;
      } else {
        changed = s(実列名.コードステータス, '入力途中') || changed;
      }
    }
  }

  if (changed) {
    sh.getRange(row, 1, 1, lastCol).setValues([rowValues]);
    if (実列名.粗利益率) {
      sh.getRange(row, 列[実列名.粗利益率]).setNumberFormat('0.0%');
    }
  }
}

function 台湾書籍系_onEdit_共通_(e, 設定) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== 設定.マスターシート名) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const col = e.range.getColumn();
  const 列 = 台湾書籍系_列マップを取得_(sh);
  const 監視列番号 = (設定.監視列 || [])
    .concat(['発番発行', '売価', '原価', 'サイト商品コード'])
    .map(name => 列[name])
    .filter(Boolean);

  if (!監視列番号.includes(col)) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;

  try {
    台湾書籍系_1行補完_共通_(sh, row, 設定);
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
      台湾書籍系_1行補完_共通_(sh, r, 設定);
    }
  } finally {
    lock.releaseLock();
  }
}

/* --- シート別ラッパー --- */
function 台湾まんが_onEdit(e) {
  台湾書籍系_onEdit_共通_(e, 設定_台湾まんが);
}

function 台湾書籍その他_onEdit(e) {
  台湾書籍系_onEdit_共通_(e, 設定_台湾書籍その他);
}

function 台湾まんが_取込後補完_(startRow, numRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾まんが');
  台湾書籍系_追加行補完_共通_(sh, startRow, numRows, 設定_台湾まんが);
}

function 台湾書籍その他_取込後補完_(startRow, numRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾書籍その他');
  台湾書籍系_追加行補完_共通_(sh, startRow, numRows, 設定_台湾書籍その他);
}

