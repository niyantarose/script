/**
 * 共通.gs
 * 全シート共通のユーティリティ・Works管理・SKU/タイトル生成
 *
 * ★ 設計方針
 * - シート固有の設定は各シートファイルの cfg オブジェクトに持つ
 * - 共通関数は cfg を引数で受け取り、グローバル変数に依存しない
 * - cfg の構造: { マスターシート名, 作品シート名, 言語マスター名, カテゴリマスター名,
 *                 形態マスター名, 作品ヘッダー, 作品列数, 色パレット,
 *                 列名: { ... } }
 */

/* ============================================================
 * ヘッダー正規化・列番号取得
 * ============================================================ */
function ヘッダー正規化_(s) {
  return String(s || '')
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/／/g, '/').replace(/　/g, ' ')
    .replace(/\s+/g, ' ').trim();
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

/* ============================================================
 * WorksKey生成（原題優先）
 * ============================================================ */
function WorksKeyを作る(日本語タイトル, 作者, 原題 = '') {
  const 正規化原題 = キー用正規化_(原題);
  if (正規化原題) return '原題||' + 正規化原題;
  return キー用正規化_(日本語タイトル) + '||' + キー用正規化_(作者);
}

/* ============================================================
 * Works照合と同名疑い警告
 * ============================================================ */
function Works照合と警告(worksKey, 作者, 作品データ) {
  const result = { 作品ID: null, 警告: '' };
  const 既存ID = 作品データ.keyToId[worksKey];
  if (!既存ID) return result;
  result.作品ID = String(既存ID).padStart(4, '0');
  if (worksKey.startsWith('原題||') && 作者) {
    const 既存作者 = 作品データ.keyToData[worksKey]?.作者 || '';
    if (既存作者 && キー用正規化_(作者) !== キー用正規化_(既存作者)) {
      result.警告 = '【同名疑い・要確認】';
    }
  }
  return result;
}

/* ============================================================
 * 商品行重複チェックキー生成
 * ============================================================ */
function 商品行キーを作る(原題, 単巻数, セット開始, セット終了, 形態, 言語) {
  const タイトル部 = キー用正規化_(原題) || 'NOTITLE';
  const 形態部 = キー用正規化_(形態) || '通常';
  const 言語部 = String(言語 || '').trim();
  const sf = String(セット開始 || '').trim();
  const st = String(セット終了 || '').trim();
  let 巻数部;
  if (sf && st) {
    巻数部 = `${sf}-${st}`;
  } else {
    const v = 数値変換(単巻数);
    巻数部 = v != null ? String(v).padStart(2, '0') : 'NOVOL';
  }
  return `${タイトル部}||${巻数部}||${形態部}||${言語部}`;
}

/* ============================================================
 * onEdit ディスパッチ（全シート共通エントリーポイント）
 * ============================================================ */
function メインonEdit(e) {  // onEdit → メインonEdit に変更
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const shName = sh.getName();

  const ALADIN_TARGET_SHEETS = ['韓国書籍', '韓国マンガ', '韓国音楽映像'];
  const ALADIN_TRIGGERS = ['アラジンURL', 'ISBN', 'JANコード'];

  if (ALADIN_TARGET_SHEETS.includes(shName)) {
    アラジン_onEdit(e);
    // アラジンURL/ISBN入力時は共通フレームワークをスキップ（ループ防止）
    const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const 編集列名 = String(ヘッダー[e.range.getColumn() - 1] || '').trim();
    if (ALADIN_TRIGGERS.includes(編集列名)) return;
  }

  if (shName === '台湾グッズ') { 台湾グッズ_onEdit(e); return; }
  if (shName === '韓国グッズ') { アラジン_onEdit(e); 韓国グッズ_onEdit(e); return; }
  if (shName === '中国価格計算') { 中国価格_onEdit(e); return; }

  const cfg = シート設定を取得(shName);
  if (!cfg) return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  if (開始行 + 行数 - 1 < 2) return;

  const lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(5000)) return; } catch (_) { return; }

  try {
    if (自己更新中か_()) return;
    const 列マップ = 列番号を取得(sh);
    const 編集開始列 = e.range.getColumn();
    const 編集終了列 = e.range.getLastColumn();
    const 監視列番号 = cfg.監視列.map(h => 列マップ[h]).filter(Boolean);
    const 対象列が含まれる = 監視列番号.some(c => c >= 編集開始列 && c <= 編集終了列);
    if (!対象列が含まれる) return;
    自己更新を開始_();
    try {
      onEdit処理を実行(e, sh, cfg, 列マップ, 開始行, 行数);
    } finally {
      自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}
function onEdit処理を実行(e, sh, cfg, 列マップ, 開始行, 行数) {
  const 最終列 = sh.getLastColumn();
  const 行データ一覧 = sh.getRange(開始行, 1, 行数, 最終列).getValues();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = 作品シートを確保(ss, cfg);
  const 作品データ = 全作品データを読み込み(作品シート, cfg);
  const 言語マップ = 言語マップを取得(ss, cfg);
  const カテゴリマップ = カテゴリマップを取得(ss, cfg);
  const 形態マップ = 形態マップを取得(ss, cfg);

  const 出力SKU = [], 出力タイトル = [], 出力作品ID = [], 出力SKU自動 = [], 出力ステータス = [];
  const ISBNトラッカー = {}, 商品行キートラッカー = {}, 予約巻数マップ = {};
  const cn = cfg.列名;

  for (let i = 0; i < 行数; i++) {
    const 行番号 = 開始行 + i;
    const 行データ = 行データ一覧[i];
    if (行番号 < 2) {
      出力SKU.push(['']); 出力タイトル.push(['']); 出力作品ID.push(['']);
      出力SKU自動.push(['']); 出力ステータス.push(['']);
      continue;
    }

    const 取得 = (名前) => 正規化(行データ[(列マップ[名前] || 1) - 1]);
    const 取得生 = (名前) => 行データ[(列マップ[名前] || 1) - 1];

    const 日本語タイトル = 取得(cn.日本語タイトル);
    let 作者 = 取得(cn.作者);
    let 原題 = 取得(cn.原題);
    const 言語 = 取得(cn.言語);
    const カテゴリ = 取得(cn.カテゴリ);
    const 形態 = 取得(cn.形態);
    const 単巻数 = 取得生(cn.単巻数);
    const セット開始 = 取得生(cn.セット開始);
    const セット終了 = 取得生(cn.セット終了);
    const 特典メモ = 取得(cn.特典メモ);
    const ISBN = 取得(cn.ISBN);

    if (!日本語タイトル || !作者) {
      出力SKU.push([取得生(cn.商品コード) || '']);
      出力タイトル.push([取得生(cn.タイトル) || '']);
      出力作品ID.push([取得生(cn.作品ID) || '']);
      出力SKU自動.push([取得生(cn.SKU自動) || '']);
      出力ステータス.push([取得生(cn.コードステータス) || '']);
      continue;
    }

    const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
    const 全条件揃い = !!(日本語タイトル && 作者 && 言語 && カテゴリ);
    let 作品ID = '', 同名警告 = '';

    const 照合結果 = Works照合と警告(worksKey, 作者, 作品データ);
    if (照合結果.作品ID) {
      作品ID = 照合結果.作品ID;
      同名警告 = 照合結果.警告;
      const 既存 = 作品データ.keyToData[worksKey];
      if (既存) {
        if (!作者 && 既存.作者) 作者 = 既存.作者;
        if (!原題 && 既存.原題) 原題 = 既存.原題;
      }
    } else if (全条件揃い) {
      if (!worksKey.startsWith('原題||')) {
        const titleKey = キー用正規化_(日本語タイトル);
        const 既存worksKey = 作品データ.titleToKey[titleKey];
        if (既存worksKey && 作品データ.keyToId[既存worksKey]) {
          作品ID = String(作品データ.keyToId[既存worksKey]).padStart(4, '0');
          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.keyToRow[worksKey] = 作品データ.keyToRow[既存worksKey];
          作品データ.keyToData[worksKey] = 作品データ.keyToData[既存worksKey];
        }
      }
      if (!作品ID) {
        作品データ.maxId++;
        作品ID = String(作品データ.maxId).padStart(4, '0');
        作品データ.keyToId[worksKey] = 作品ID;
        作品データ.keyToRow[worksKey] = null;
        作品データ.keyToData[worksKey] = { 作者, 原題 };
        作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題を正規化(原題), '', '', '', '', '']);
        if (!worksKey.startsWith('原題||')) {
          const titleKey = キー用正規化_(日本語タイトル);
          if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey;
        }
      }
    } else {
      作品ID = '????';
    }

    // 重複チェック
    let 巻数警告 = '';
    if (ISBN) {
      if (ISBNトラッカー[ISBN]) 巻数警告 = '【重複注意】';
      ISBNトラッカー[ISBN] = true;
    } else if (原題) {
      const 商品行キー = 商品行キーを作る(原題, 単巻数, セット開始, セット終了, 形態, 言語);
      if (商品行キートラッカー[商品行キー]) 巻数警告 = '【重複注意】';
      商品行キートラッカー[商品行キー] = true;
    }

    // 予約巻数マップ
    const 巻数 = 数値変換(単巻数);
    const セット終了巻 = 数値変換(セット終了);
    const 最新巻候補 = 巻数 != null ? 巻数 : セット終了巻;
    if (最新巻候補 != null && 作品ID !== '????') {
      if (!予約巻数マップ[worksKey]) 予約巻数マップ[worksKey] = [];
      予約巻数マップ[worksKey].push(最新巻候補);
    }

    const 言語コード = 言語 ? (言語マップ[言語] || 'XX') : '';
    const カテゴリコード = カテゴリ ? (カテゴリマップ[カテゴリ] || 'XX') : '';
    const 形態コード = 形態 ? (形態マップ.コード[形態] || '') : '';
    const SKU = SKUを段階生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了);
    const タイトル = 同名警告 + 巻数警告 + タイトルを段階生成(cfg, 言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ, 形態マップ);

    let ステータス = '';
    if (同名警告) ステータス = 同名警告;
    else if (巻数警告) ステータス = '【重複注意】';
    else ステータス = 全条件揃い ? '商品コード(予約)' : '入力中...';

    出力SKU.push([SKU]);
    出力タイトル.push([タイトル]);
    出力作品ID.push([作品ID]);
    出力SKU自動.push([SKU]);
    出力ステータス.push([ステータス]);
  }

  if (列マップ[cn.商品コード])     sh.getRange(開始行, 列マップ[cn.商品コード],     行数, 1).setValues(出力SKU);
  if (列マップ[cn.タイトル])       sh.getRange(開始行, 列マップ[cn.タイトル],       行数, 1).setValues(出力タイトル);
  if (列マップ[cn.作品ID])         sh.getRange(開始行, 列マップ[cn.作品ID],         行数, 1).setValues(出力作品ID);
  if (列マップ[cn.SKU自動])        sh.getRange(開始行, 列マップ[cn.SKU自動],        行数, 1).setValues(出力SKU自動);
  if (列マップ[cn.コードステータス]) sh.getRange(開始行, 列マップ[cn.コードステータス], 行数, 1).setValues(出力ステータス);

  作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ, cfg);
}

/* ============================================================
 * Works更新（onEdit専用: 新規登録 + 予約I/J列更新）
 * ============================================================ */
function 作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ, cfg) {
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    const 最終行 = 作品シート.getLastRow();
    if (最終行 >= 2) {
      const 全Works = 作品シート.getRange(2, 1, 最終行 - 1, 1).getValues();
      for (const upd of 作品データ.keyUpdates) 全Works[upd.行 - 2][0] = upd.key;
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
 * Works更新（確定版: ①確定発行・③一括更新で使用）
 * ============================================================ */
function 作品データを更新_確定(作品シート, 作品データ, cfg) {
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
  if (!作品シート || 作品シート.getLastRow() < 2) return { キー更新数: 0, 統合数: 0, 削除行数: 0 };
  const 最終行 = 作品シート.getLastRow();
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, cfg.作品列数).getValues();

  // 前処理: 同タイトル+作者グループで原題を統合
  const タイトル作者グループ = new Map();
  for (let i = 0; i < データ.length; i++) {
    const t = 正規化(データ[i][2] || ''), a = 正規化(データ[i][3] || '');
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
    const t = 正規化(r[2] || ''), a = 正規化(r[3] || ''), o = 正規化(r[4] || '');
    if (!t && !a) continue;
    const key = WorksKeyを作る(t, a, o);
    const 元key = String(r[0] || '').trim();
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: r, 元key, キー変更: 元key !== key });
  }

  let キー更新数 = 0, 統合数 = 0;
  const 削除行 = [];
  const 更新データ = データ.map(r => r.slice());

  for (const [key, 行配列] of keyMap.entries()) {
    if (行配列.length === 1) {
      if (行配列[0].キー変更) { 更新データ[行配列[0].行 - 2][0] = key; キー更新数++; }
    } else {
      行配列.sort((a, b) => {
        const a原題 = 正規化(a.データ[4] || '') ? 0 : 1;
        const b原題 = 正規化(b.データ[4] || '') ? 0 : 1;
        if (a原題 !== b原題) return a原題 - b原題;
        return b.行 - a.行;
      });
      const 残す = 行配列[0];
      更新データ[残す.行 - 2][0] = key;
      const 全巻 = new Set();
      for (const item of 行配列) {
        String(item.データ[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) 全巻.add(n); });
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
      if (最大予約巻 > 0) { 更新データ[残す.行 - 2][8] = 最大予約巻; 更新データ[残す.行 - 2][9] = new Date(); }
      if (!正規化(残す.データ[4] || '')) {
        for (let j = 1; j < 行配列.length; j++) {
          const t = 正規化(行配列[j].データ[4] || '');
          if (t) { 更新データ[残す.行 - 2][4] = t; break; }
        }
      }
      for (let j = 1; j < 行配列.length; j++) 削除行.push(行配列[j].行);
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
  const result = { 変更数: 0, 旧新マップ: {} };
  if (!作品シート || 作品シート.getLastRow() < 2) return result;
  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, cfg.作品列数).getValues();
  const 行リスト = データ.map(r => ({ 旧ID: parseInt(String(r[1] || '0'), 10), データ: r }));
  行リスト.sort((a, b) => a.旧ID - b.旧ID);
  const 新データ = 行リスト.map((item, i) => {
    const r = item.データ.slice();
    const 新ID = String(i + 1).padStart(4, '0');
    if (String(item.旧ID).padStart(4, '0') !== 新ID) { result.旧新マップ[String(item.旧ID).padStart(4, '0')] = 新ID; result.変更数++; }
    r[1] = 新ID;
    return r;
  });
  作品シート.getRange(2, 1, 新データ.length, cfg.作品列数).setValues(新データ);
  return result;
}

/* ============================================================
 * Works重複検出
 * ============================================================ */
function Works重複を検出(作品シート, cfg) {
  if (!作品シート || 作品シート.getLastRow() < 2) return [];
  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, cfg.作品列数).getValues();
  const keyMap = new Map();
  for (let i = 0; i < データ.length; i++) {
    const t = 正規化(データ[i][2]), a = 正規化(データ[i][3]), o = 正規化(データ[i][4]);
    if (!t || !a) continue;
    const key = WorksKeyを作る(t, a, o);
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: データ[i] });
  }
  return Array.from(keyMap.entries()).filter(([_, v]) => v.length > 1).map(([key, 行配列]) => ({ key, 行配列 }));
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
      const t = 正規化(r[(列マップ[cn.日本語タイトル] || 1) - 1]);
      const a = 正規化(r[(列マップ[cn.作者] || 1) - 1]);
      const o = 正規化(r[(列マップ[cn.原題] || 1) - 1]);
      if (t && a) {
        使用中Keys.add(WorksKeyを作る(t, a, o));
        使用中Keys.add(キー用正規化_(t) + '||' + キー用正規化_(a));
      }
    }
  }
  const worksData = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, cfg.作品列数).getValues();
  const 削除行 = [];
  for (let i = 0; i < worksData.length; i++) {
    const t = 正規化(worksData[i][2] || ''), a = 正規化(worksData[i][3] || ''), o = 正規化(worksData[i][4] || '');
    if (!t && !a) { 削除行.push(i + 2); continue; }
    const computedKey = WorksKeyを作る(t, a, o);
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
      for (let i = 現在列数; i < cfg.作品列数; i++) sh.getRange(1, i + 1).setValue(cfg.作品ヘッダー[i]);
    }
  }
  return sh;
}

function 全作品データを読み込み(作品シート, cfg) {
  const result = {
    keyToId: {}, keyToData: {}, keyToRow: {}, keyToVols: {},
    keyTo予約最新巻: {}, maxId: 0, newRows: [], keyUpdates: [], titleToKey: {}
  };
  const 最終行 = 作品シート.getLastRow();
  if (最終行 < 2) return result;
  const 列数 = Math.max(作品シート.getLastColumn(), cfg.作品列数);
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, Math.min(列数, cfg.作品列数)).getValues();
  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const idStr = String(r[1] == null ? '' : r[1]).trim();
    if (!idStr) continue;
    const t = 正規化(r[2] || ''), a = 正規化(r[3] || ''), o = 正規化(r[4] || '');
    let key = (t && a) ? WorksKeyを作る(t, a, o) : String(r[0] || '').trim();
    if (!key) continue;
    const 保存済みKey = String(r[0] || '').trim();
    if (保存済みKey !== key) result.keyUpdates.push({ 行: i + 2, key });
    if (result.keyToId[key]) {
      const 既存ID = parseInt(result.keyToId[key], 10), 新ID = parseInt(idStr, 10);
      if (!isNaN(新ID) && !isNaN(既存ID) && 新ID >= 既存ID) {
        String(r[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) { if (!result.keyToVols[key]) result.keyToVols[key] = new Set(); result.keyToVols[key].add(n); } });
        const num = parseInt(idStr, 10); if (!isNaN(num) && num > result.maxId) result.maxId = num;
        continue;
      }
    }
    result.keyToId[key] = idStr.padStart(4, '0');
    result.keyToRow[key] = i + 2;
    result.keyToData[key] = { 作者: 正規化(r[3] || ''), 原題: 正規化(r[4] || '') };
    if (!result.keyToVols[key]) result.keyToVols[key] = new Set();
    String(r[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) result.keyToVols[key].add(n); });
    const 予約巻 = parseInt(String(r[8] || '0'), 10);
    if (!isNaN(予約巻) && 予約巻 > 0) result.keyTo予約最新巻[key] = 予約巻;
    const num = parseInt(idStr, 10); if (!isNaN(num) && num > result.maxId) result.maxId = num;
    if (t && !key.startsWith('原題||')) {
      const titleKey = キー用正規化_(t);
      if (!result.titleToKey[titleKey]) result.titleToKey[titleKey] = key;
    }
  }
  return result;
}

/* ============================================================
 * ① 確定発行（共通実装）
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
  const 言語マップ = 言語マップを取得(ss, cfg);
  const カテゴリマップ = カテゴリマップを取得(ss, cfg);
  const 形態マップ = 形態マップを取得(ss, cfg);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss, cfg);
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    WorksKey再正規化を実行(作品シート, cfg);
    const 作品データ = 全作品データを読み込み(作品シート, cfg);
    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [], outステータス = [];
    let 発行数 = 0;
    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const 取得 = (名前) => 正規化(r[(列マップ[名前] || 1) - 1]);
      const 取得生 = (名前) => r[(列マップ[名前] || 1) - 1];
      if (取得生(cn.発行チェック) !== true) {
        out作品ID.push([取得生(cn.作品ID) || '']); outSKU.push([取得生(cn.SKU自動) || '']);
        out商品コード.push([取得生(cn.商品コード) || '']); outタイトル.push([取得生(cn.タイトル) || '']);
        outステータス.push([取得生(cn.コードステータス) || '']);
        continue;
      }
      const 日本語タイトル = 取得(cn.日本語タイトル), 作者 = 取得(cn.作者);
      let 原題 = 取得(cn.原題);
      const 言語 = 取得(cn.言語), カテゴリ = 取得(cn.カテゴリ), 形態 = 取得(cn.形態);
      if (!日本語タイトル || !作者 || !言語 || !カテゴリ) {
        out作品ID.push([取得生(cn.作品ID) || '']); outSKU.push([取得生(cn.SKU自動) || '']);
        out商品コード.push([取得生(cn.商品コード) || '']); outタイトル.push([取得生(cn.タイトル) || '']);
        outステータス.push(['入力中...']); continue;
      }
      const 言語コード = 言語マップ[言語] || 'XX', カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const 形態コード = 形態 ? (形態マップ.コード[形態] || '') : '';
      const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
      let 作品ID = 作品データ.keyToId[worksKey];
      const 既存 = 作品データ.keyToData[worksKey];
      if (既存 && !原題 && 既存.原題) 原題 = 正規化(既存.原題);
      if (!作品ID) {
        if (!worksKey.startsWith('原題||')) {
          const titleKey = キー用正規化_(日本語タイトル);
          const 既存worksKey = 作品データ.titleToKey[titleKey];
          if (既存worksKey && 作品データ.keyToId[既存worksKey]) { 作品ID = 作品データ.keyToId[既存worksKey]; 作品データ.keyToId[worksKey] = 作品ID; }
        }
        if (!作品ID) {
          作品データ.maxId++;
          作品ID = String(作品データ.maxId).padStart(4, '0');
          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題を正規化(原題), '', '', '', '', '']);
          if (!worksKey.startsWith('原題||')) { const titleKey = キー用正規化_(日本語タイトル); if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey; }
        }
      } else { 作品ID = String(作品ID).padStart(4, '0'); }

      const 巻数 = 数値変換(取得生(cn.単巻数));
      const セット終了巻 = 数値変換(取得生(cn.セット終了));
      if (巻数 != null) { if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set(); 作品データ.keyToVols[worksKey].add(巻数); }
      if (セット終了巻 != null) {
        const セット開始巻 = 数値変換(取得生(cn.セット開始)) || 1;
        if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
        for (let v = セット開始巻; v <= セット終了巻; v++) 作品データ.keyToVols[worksKey].add(v);
      }
      const SKU = SKUを生成(言語コード, 形態コード, 作品ID, カテゴリコード, 取得生(cn.単巻数), 取得生(cn.セット開始), 取得生(cn.セット終了));
      const タイトル = タイトルを段階生成(cfg, 言語, カテゴリ, 形態, 日本語タイトル, 取得生(cn.単巻数), 取得生(cn.セット開始), 取得生(cn.セット終了), 作者, 原題, 取得(cn.特典メモ), 形態マップ);
      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]);
      outタイトル.push([タイトル]); outステータス.push(['商品コード(発行済み確定)']);
      発行数++;
    }
    if (列マップ[cn.作品ID])         sh.getRange(2, 列マップ[cn.作品ID],         out作品ID.length,    1).setValues(out作品ID);
    if (列マップ[cn.SKU自動])        sh.getRange(2, 列マップ[cn.SKU自動],        outSKU.length,       1).setValues(outSKU);
    if (列マップ[cn.商品コード])     sh.getRange(2, 列マップ[cn.商品コード],     out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[cn.タイトル])       sh.getRange(2, 列マップ[cn.タイトル],       outタイトル.length,   1).setValues(outタイトル);
    if (列マップ[cn.コードステータス]) sh.getRange(2, 列マップ[cn.コードステータス], outステータス.length, 1).setValues(outステータス);
    作品データを更新_確定(作品シート, 作品データ, cfg);
    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * ② 削除（共通実装）
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
    if (全データ[i][(列マップ[cn.発行チェック] || 1) - 1] === true) 削除行.push(i + 2);
  }
  if (削除行.length === 0) { ui.alert('チェックが入った行がありません'); return; }
  if (ui.alert('確認', `${削除行.length}件を削除します。続行？`, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  行を一括削除(sh, 削除行);
  ui.alert(`✅ 削除完了: ${削除行.length}件`);
}

/* ============================================================
 * ③ 一括更新（共通実装）
 * ============================================================ */
function 一括更新を実行(cfg) {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '未登録行を再生成します（Works重複整理+ID振り直し含む）。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.マスターシート名);
  if (!sh) return;
  const cn = cfg.列名;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ, cn);
  if (最終行 < 2) { ui.alert('データがありません'); return; }
  const 言語マップ = 言語マップを取得(ss, cfg);
  const カテゴリマップ = カテゴリマップを取得(ss, cfg);
  const 形態マップ = 形態マップを取得(ss, cfg);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss, cfg);
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    const 正規化結果 = WorksKey再正規化を実行(作品シート, cfg);
    const 孤立結果 = { 削除数: 0 };
    const ID結果 = WorksID振り直しを実行(作品シート, cfg);
    const 作品データ = 全作品データを読み込み(作品シート, cfg);
    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [], outステータス = [], out作者 = [], out原題 = [];

    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const 取得 = (名前) => 正規化(r[(列マップ[名前] || 1) - 1]);
      const 取得生 = (名前) => r[(列マップ[名前] || 1) - 1];
      const 日本語タイトル = 取得(cn.日本語タイトル);
      const 言語 = 取得(cn.言語), カテゴリ = 取得(cn.カテゴリ);
      let 作者 = 取得(cn.作者), 原題 = 取得(cn.原題);
      const 形態 = 取得(cn.形態);
      const ステータス = 取得生(cn.コードステータス);
      const 確定行 = ステータス === '商品コード(発行済み確定)';

      if (確定行) {
        if (日本語タイトル && 作者) {
          const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
          const 既存 = 作品データ.keyToData[worksKey];
          if (既存 && !原題 && 既存.原題) 原題 = 正規化(既存.原題);
          const 巻数 = 数値変換(取得生(cn.単巻数));
          const セット終了巻 = 数値変換(取得生(cn.セット終了));
          if (巻数 != null) { if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set(); 作品データ.keyToVols[worksKey].add(巻数); }
          if (セット終了巻 != null) {
            const セット開始巻 = 数値変換(取得生(cn.セット開始)) || 1;
            if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
            for (let v = セット開始巻; v <= セット終了巻; v++) 作品データ.keyToVols[worksKey].add(v);
          }
        }
        out作品ID.push([取得生(cn.作品ID) || '']); outSKU.push([取得生(cn.SKU自動) || '']);
        out商品コード.push([取得生(cn.商品コード) || '']); outタイトル.push([取得生(cn.タイトル) || '']);
        outステータス.push([ステータス]); out作者.push([作者 || '']); out原題.push([原題 || '']);
        continue;
      }

      if (!日本語タイトル || !言語 || !カテゴリ) {
        out作品ID.push([取得生(cn.作品ID) || '']); outSKU.push([取得生(cn.SKU自動) || '']);
        out商品コード.push([取得生(cn.商品コード) || '']); outタイトル.push([取得生(cn.タイトル) || '']);
        outステータス.push([ステータス || '']); out作者.push([作者 || '']); out原題.push([原題 || '']);
        continue;
      }

      const 言語コード = 言語マップ[言語] || 'XX', カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const 形態コード = 形態 ? (形態マップ.コード[形態] || '') : '';
      let 作品ID = '';
      if (日本語タイトル && 作者) {
        const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
        const 既存 = 作品データ.keyToData[worksKey];
        if (既存) { if (!作者 && 既存.作者) 作者 = 正規化(既存.作者); if (!原題 && 既存.原題) 原題 = 正規化(既存.原題); }
        作品ID = 作品データ.keyToId[worksKey];
        if (!作品ID) {
          if (!worksKey.startsWith('原題||')) {
            const titleKey = キー用正規化_(日本語タイトル);
            const 既存worksKey = 作品データ.titleToKey[titleKey];
            if (既存worksKey && 作品データ.keyToId[既存worksKey]) { 作品ID = 作品データ.keyToId[既存worksKey]; 作品データ.keyToId[worksKey] = 作品ID; }
          }
          if (!作品ID) {
            作品データ.maxId++;
            作品ID = String(作品データ.maxId).padStart(4, '0');
            作品データ.keyToId[worksKey] = 作品ID;
            作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題を正規化(原題), '', '', '', '', '']);
            if (!worksKey.startsWith('原題||')) { const titleKey = キー用正規化_(日本語タイトル); if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey; }
          }
        } else { 作品ID = String(作品ID).padStart(4, '0'); }
      }

      const SKU = SKUを段階生成(言語コード, 形態コード, 作品ID, カテゴリコード, 取得生(cn.単巻数), 取得生(cn.セット開始), 取得生(cn.セット終了));
      const タイトル = タイトルを段階生成(cfg, 言語, カテゴリ, 形態, 日本語タイトル, 取得生(cn.単巻数), 取得生(cn.セット開始), 取得生(cn.セット終了), 作者, 原題, 取得(cn.特典メモ), 形態マップ);
      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]); outタイトル.push([タイトル]);
      outステータス.push([(日本語タイトル && 作者 && 言語 && カテゴリ) ? '商品コード(予約)' : '入力中...']);
      out作者.push([作者 || '']); out原題.push([原題 || '']);
    }

    if (列マップ[cn.作品ID])           sh.getRange(2, 列マップ[cn.作品ID],           out作品ID.length,    1).setValues(out作品ID);
    if (列マップ[cn.SKU自動])          sh.getRange(2, 列マップ[cn.SKU自動],          outSKU.length,       1).setValues(outSKU);
    if (列マップ[cn.商品コード])       sh.getRange(2, 列マップ[cn.商品コード],       out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[cn.タイトル])         sh.getRange(2, 列マップ[cn.タイトル],         outタイトル.length,   1).setValues(outタイトル);
    if (列マップ[cn.コードステータス]) sh.getRange(2, 列マップ[cn.コードステータス], outステータス.length, 1).setValues(outステータス);
    if (列マップ[cn.作者])             sh.getRange(2, 列マップ[cn.作者],             out作者.length,       1).setValues(out作者);
    if (列マップ[cn.原題])             sh.getRange(2, 列マップ[cn.原題],             out原題.length,       1).setValues(out原題);
    作品データを更新_確定(作品シート, 作品データ, cfg);

    let msg = `✅ 一括更新完了: ${out作品ID.length}件`;
    if (正規化結果.統合数 > 0) msg += `\nWorks重複統合: ${正規化結果.統合数}件`;
    if (孤立結果.削除数 > 0) msg += `\nWorks孤立削除: ${孤立結果.削除数}件`;
    if (ID結果.変更数 > 0) msg += `\nID振り直し: ${ID結果.変更数}件`;
    ui.alert(msg);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * SKU / タイトル生成
 * ============================================================ */
function SKUを段階生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const parts = [];
  if (言語コード) parts.push(言語コード);
  if (形態コード) parts.push(形態コード);
  if (作品ID) { parts.push(String(作品ID).padStart(4, '0')); } else if (言語コード) { parts.push('????'); }
  if (カテゴリコード) parts.push('-' + カテゴリコード);
  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) {
    const diff = parseInt(st, 10) - parseInt(sf, 10);
    if (diff >= 2) { parts.push('-SET'); }
    else { parts.push('-' + sf.padStart(2, '0') + st.padStart(2, '0')); }
  } else {
    const v = String(単巻数 || '').trim();
    if (v) { const n = parseInt(v, 10); parts.push('-' + (!isNaN(n) ? String(n).padStart(2, '0') : v)); }
  }
  return parts.join('');
}

function SKUを生成(言語コード, 形態コード, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const base = String(言語コード || '') + String(形態コード || '') + String(作品ID || '').padStart(4, '0') + '-' + String(カテゴリコード || '');
  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) {
    const diff = parseInt(st, 10) - parseInt(sf, 10);
    if (diff >= 2) return base + '-SET';
    return base + '-' + sf.padStart(2, '0') + st.padStart(2, '0');
  }
  const v = String(単巻数 || '').trim();
  if (v) { const n = parseInt(v, 10); return base + '-' + (!isNaN(n) ? String(n).padStart(2, '0') : v); }
  return base;
}

function タイトルを段階生成(cfg, 言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ, 形態マップ) {
  const parts = [];
  const head = [言語 ? `${言語}版` : '', カテゴリ].filter(Boolean).join(' ');
  if (head) parts.push(head);
  if (形態 && 形態 !== '通常版' && 形態 !== '通常') parts.push(`(${形態})`);
  if (日本語タイトル) parts.push(`『${日本語タイトル}』`);

  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) {
    const diff = parseInt(st, 10) - parseInt(sf, 10);
    if (diff >= 2) { parts.push(`${sf.padStart(2, '0')}-${st.padStart(2, '0')}巻 SET`); }
    else { parts.push(`${sf.padStart(2, '0')}-${st.padStart(2, '0')}巻`); }
  } else {
    const v = String(単巻数 || '').trim();
    if (v) { const n = parseInt(v, 10); parts.push(`第${!isNaN(n) ? String(n).padStart(2, '0') : v}巻`); }
  }

  if (作者) parts.push(`著：${作者}`);

  if (原題) {
    let 原題巻数 = '';
    if (sf && st) { 原題巻数 = `${sf}-${st}`; }
    else { const v = String(単巻数 || '').trim(); if (v) { const n = parseInt(v, 10); 原題巻数 = !isNaN(n) ? String(n) : v; } }
    const 原題形態表記 = 形態マップ.原題形態[形態] || '';
    let 構築原題 = 原題;
    if (原題巻数) 構築原題 += 原題巻数;
    if (原題形態表記) 構築原題 += ` (${原題形態表記})`;
    parts.push(構築原題);
  }

  if (特典メモ) { parts.push(特典メモ.startsWith('※') ? 特典メモ : `※${特典メモ}`); }
  else if (形態マップ.特典テキスト[形態]) { parts.push(形態マップ.特典テキスト[形態]); }

  return parts.join(' ');
}

/* ============================================================
 * マスターデータ読み込み
 * ============================================================ */
function 言語マップを取得(ss, cfg) {
  const map = {}, sh = ss.getSheetByName(cfg.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) { const name = String(n || '').trim(), code = String(c || '').trim(); if (name && code) map[name] = code; }
  return map;
}

function カテゴリマップを取得(ss, cfg) {
  const map = {}, sh = ss.getSheetByName(cfg.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) { const name = String(n || '').trim(), code = String(c || '').trim(); if (name && code) map[name] = code; }
  return map;
}

function 形態マップを取得(ss, cfg) {
  // 形態マスターシートから { コード, 特典テキスト, 原題形態 } を読み込む
  // 列構成: A=形態名, B=コード, C=特典テキスト, D=色, E=原題形態表記(台湾用)
  const result = { コード: {}, 特典テキスト: {}, 原題形態: {} };
  const sh = ss.getSheetByName(cfg.形態マスター名 || '形態マスターシート');
  if (!sh || sh.getLastRow() < 2) return result;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (const row of data) {
    const 名前 = String(row[0] || '').trim();
    if (!名前) continue;
    result.コード[名前] = String(row[1] || '').trim();
    result.特典テキスト[名前] = String(row[2] || '').trim();
    result.原題形態[名前] = String(row[4] || '').trim();
  }
  return result;
}

function 言語データを色付きで取得(ss, cfg) {
  const result = { values: [], colors: [] }, sh = ss.getSheetByName(cfg.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()) {
    const n = String(row[0] || '').trim(); if (!n) continue;
    result.values.push(n); result.colors.push(String(row[2] || '').trim() || cfg.色パレット[ci++ % cfg.色パレット.length]);
  }
  return result;
}

function カテゴリデータを色付きで取得(ss, cfg) {
  const result = { values: [], colors: [] }, sh = ss.getSheetByName(cfg.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()) {
    const n = String(row[0] || '').trim(); if (!n) continue;
    result.values.push(n); result.colors.push(String(row[2] || '').trim() || cfg.色パレット[ci++ % cfg.色パレット.length]);
  }
  return result;
}

function 形態データを色付きで取得(ss, cfg) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName(cfg.形態マスター名 || '形態マスターシート');
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()) {
    const n = String(row[0] || '').trim(); if (!n) continue;
    result.values.push(n); result.colors.push(String(row[3] || '').trim() || cfg.色パレット[ci++ % cfg.色パレット.length]);
  }
  return result;
}

function プルダウン更新を実行(cfg) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.マスターシート名);
  if (!sh) return;
  const cn = cfg.列名;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = Math.max(データがある最終行を取得(sh, 列マップ, cn) + 100, 200);

  const 言語データ = 言語データを色付きで取得(ss, cfg);
  if (言語データ.values.length > 0 && 列マップ[cn.言語]) {
    sh.getRange(2, 列マップ[cn.言語], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(言語データ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[cn.言語]);
    条件付き書式を設定(sh, 列マップ[cn.言語], 言語データ.values, 言語データ.colors, 最終行);
  }

  const カテゴリデータ = カテゴリデータを色付きで取得(ss, cfg);
  if (カテゴリデータ.values.length > 0 && 列マップ[cn.カテゴリ]) {
    sh.getRange(2, 列マップ[cn.カテゴリ], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(カテゴリデータ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[cn.カテゴリ]);
    条件付き書式を設定(sh, 列マップ[cn.カテゴリ], カテゴリデータ.values, カテゴリデータ.colors, 最終行);
  }

  const 形態データ = 形態データを色付きで取得(ss, cfg);
  if (形態データ.values.length > 0 && 列マップ[cn.形態]) {
    sh.getRange(2, 列マップ[cn.形態], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(形態データ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[cn.形態]);
    条件付き書式を設定(sh, 列マップ[cn.形態], 形態データ.values, 形態データ.colors, 最終行);
  }

  // ショップ名プルダウン（全シートに反映）
  const ショップシート = ss.getSheetByName('ショップマスター');
  if (ショップシート) {
    const ショップ最終行 = ショップシート.getLastRow();
    if (ショップ最終行 >= 2) {
      const ショップ値 = ショップシート.getRange(2, 1, ショップ最終行 - 1, 1)
        .getValues().flat().filter(v => v !== '');
      if (ショップ値.length > 0) {
        ss.getSheets().forEach(targetSh => {
          if (targetSh.getLastColumn() === 0) return; // 空シートをスキップ
          const target列マップ = 列番号を取得(targetSh);
          const shopCol = target列マップ['ショップ名'];
          if (!shopCol || shopCol < 1) return;
          const targetRows = Math.max(targetSh.getLastRow(), 1) + 100;
          targetSh.getRange(2, shopCol, targetRows, 1)
            .setDataValidation(SpreadsheetApp.newDataValidation()
              .requireValueInList(ショップ値, true).build());
        });
      }
    }
  }

  ss.toast('プルダウン更新完了', '商品コード管理', 3);
}
/* ============================================================
 * 条件付き書式
 * ============================================================ */
function 条件付き書式をクリア(sh, colNum) {
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(rng => {
      const c1 = rng.getColumn();
      const c2 = c1 + rng.getNumColumns() - 1;
      return c1 === colNum && c2 === colNum; // 単一列ルールのみ削除
    })
  ));
}

function 条件付き書式を設定(sh, colNum, values, colors, lastRow) {
  const rules = sh.getConditionalFormatRules();
  const range = sh.getRange(2, colNum, lastRow - 1, 1);
  for (let i = 0; i < values.length; i++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(values[i]).setBackground(colors[i] || '#fff').setRanges([range]).build());
  }
  sh.setConditionalFormatRules(rules);
}

/* ============================================================
 * ヘルパー関数
 * ============================================================ */
function 正規化(v) { return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim(); }
function 数値変換(v) { const n = parseInt(String(v || '').trim(), 10); return Number.isFinite(n) ? n : null; }

function データがある最終行を取得(sh, 列マップ, cn) {
  const 基準列 = 列マップ[cn.日本語タイトル] || 1;
  const 最大行 = sh.getLastRow();
  if (最大行 <= 1) return 1;
  const data = sh.getRange(2, 基準列, 最大行 - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) { if (data[i][0] !== '' && data[i][0] !== null && data[i][0] !== undefined) return i + 2; }
  return 1;
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
    .replace(/\s*[\/／]\s*/g, ' / ').replace(/\s+/g, ' ').replace(/-+/g, '-').trim();
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

function 行を一括削除(sh, 削除行番号リスト) {
  if (削除行番号リスト.length === 0) return;
  const sorted = [...削除行番号リスト].sort((a, b) => b - a);
  let i = 0;
  while (i < sorted.length) {
    const start = sorted[i];
    let count = 1;
    while (i + count < sorted.length && sorted[i + count] === start - count) count++;
    sh.deleteRows(start - count + 1, count);
    i += count;
  }
}

/* ============================================================
 * 自己書き込みループ防止
 * ============================================================ */
function 自己更新中か_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('__SELF_EDIT_LOCK__') !== '1') return false;
  const ts = Number(props.getProperty('__SELF_EDIT_TS__') || '0');
  if (ts > 0 && (Date.now() - ts) > 10000) { props.deleteProperty('__SELF_EDIT_LOCK__'); props.deleteProperty('__SELF_EDIT_TS__'); return false; }
  return true;
}
function 自己更新を開始_() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('__SELF_EDIT_LOCK__', '1');
  props.setProperty('__SELF_EDIT_TS__', String(Date.now()));
}
function 自己更新を終了_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('__SELF_EDIT_LOCK__');
  props.deleteProperty('__SELF_EDIT_TS__');
}
