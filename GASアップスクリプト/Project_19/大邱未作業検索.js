// ============================================================
//  発注リスト大邱データ → 大邱未作業データ（色なし行の抽出＋多角度検索＋EMS大邱へ送る）
//
//  ・発注リスト大邱データで行に背景色（オレンジ/黄色=EMS積載済み）が
//    付いていない行だけを「大邱未作業データ」シートに色なしで書き出す
//    （手塗りの背景色に加えて、消込の条件付き書式と同じ条件
//      「入荷数あり×数量あり×残り列に値」も積載済みとして除外）
//  ・A列: チェックボックス。チェックしてメニューからEMS大邱へ送れる
//  ・B1: 全列キーワード検索（空白・改行を除去して部分一致）
//  ・2行目（水色）: 列ごとの絞り込み（スペース区切りAND・部分一致）
//  ・検索セルを編集すると onEdit 経由で自動絞り込み
//  ・列の並びは発注リスト大邱データと同じ（B=発注日 … S=支払金額）
// ============================================================

const MISAGYO_CFG = {
  SRC_SHEET: '発注リスト大邱データ',
  DST_SHEET: '大邱未作業データ',

  SRC_HEADER_JP: 4,    // 日本語ヘッダー行
  SRC_HEADER_EN: 5,    // 英語ヘッダー行
  SRC_DATA_START: 6,   // データ開始行
  SRC_COL_START: 2,    // B列から
  SRC_COL_END: 19,     // S列（支払金額）まで
  COLOR_CHECK_COL: 2,  // 色判定はB列（発注日）の背景色
  SRC_COL_SOSHIN: 25,  // Y列 EMS送信済（EMS大邱作業データへ送った数量。自動再計算が記入）

  DST_CHECK_COL: 1,    // A列 チェックボックス
  DST_KEYWORD_ROW: 1,  // B1 = 全列キーワード
  DST_FILTER_ROW: 2,   // 列別絞り込み行
  DST_HEADER_JP: 3,
  DST_HEADER_EN: 4,
  DST_DATA_START: 5,

  FONT_SIZE: 13,       // 文字の大きさ（発注リスト大邱データに合わせる）
  ROW_HEIGHT: 27       // 行の高さ
};

// ---------- メニュー ----------
function 大邱未作業メニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('大邱未作業')
    .addItem('🔄 未作業データを更新', '大邱未作業_更新')
    .addItem('🔍 絞り込みを実行', '大邱未作業_絞り込み')
    .addItem('🧹 検索条件をクリア', '大邱未作業_検索条件をクリア')
    .addSeparator()
    .addItem('📦 チェック行：入荷数を発注数量にする', '大邱未作業_チェック行の入荷数を発注数量にする')
    .addItem('📦 チェック行をEMS大邱へ送る', '大邱未作業_チェック行をEMS大邱へ送る')
    .addItem('📥 入荷・オプション・weightを発注リスト大邱へ同期', '大邱未作業_入荷を発注リスト大邱へ同期')
    .addToUi();
}

// ---------- メイン：未作業データを再構築 ----------
function 大邱未作業_更新() {
  const result = 大邱未作業_再構築_(true);
  if (result) {
    SpreadsheetApp.getActive().toast(
      '未作業 ' + result.total + '件を書き出しました。',
      '大邱未作業データ', 5
    );
  }
}

function 大邱未作業_絞り込み() {
  const result = 大邱未作業_再構築_(false);
  if (result) {
    SpreadsheetApp.getActive().toast(
      '該当 ' + result.shown + '件 / 未作業 ' + result.total + '件',
      '大邱未作業データ', 5
    );
  }
}

function 大邱未作業_検索条件をクリア() {
  const cfg = MISAGYO_CFG;
  const sh = SpreadsheetApp.getActive().getSheetByName(cfg.DST_SHEET);
  if (sh) {
    const width = cfg.SRC_COL_END - cfg.SRC_COL_START + 1;
    sh.getRange(cfg.DST_KEYWORD_ROW, 2).clearContent();          // B1
    sh.getRange(cfg.DST_FILTER_ROW, 2, 1, width).clearContent(); // 2行目 B..S
  }
  大邱未作業_更新();
}

// ---------- 自動同期（発注リスト大邱データ → 大邱未作業データ） ----------
// 発注リスト大邱データの編集・色塗り・行追加のたびに再構築すると編集が重くなるので、
// そのときは「要同期」フラグを立てるだけ（瞬時）。
// 大邱未作業データのシートを開いた瞬間（onSelectionChange）に、
// フラグが立っていれば自動で再構築する＝見るときは常に最新。
function 大邱未作業_同期予約_() {
  try {
    CacheService.getDocumentCache().put('MISAGYO_DIRTY', '1', 21600); // 6時間有効
  } catch (e) {}
}

function 大邱未作業_同期解除_() {
  try {
    CacheService.getDocumentCache().remove('MISAGYO_DIRTY');
  } catch (e) {}
}

function 大邱未作業_同期が必要_() {
  try {
    return CacheService.getDocumentCache().get('MISAGYO_DIRTY') === '1';
  } catch (e) {
    return false;
  }
}

// 統合onSelectionChangeから呼ばれる: 大邱未作業データを開いたら必要なときだけ最新化
function 大邱未作業_onSelectionChange_(e) {
  if (!大邱未作業_同期が必要_()) return;

  const cache = CacheService.getDocumentCache();
  if (cache.get('MISAGYO_SYNCING') === '1') return; // 二重起動防止
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(500)) return; // 混んでいたら次の選択変更でまた試す
  try {
    cache.put('MISAGYO_SYNCING', '1', 60);
    SpreadsheetApp.getActive().toast('発注リスト大邱データの変更を反映しています…', '大邱未作業データ 自動同期', 3);
    大邱未作業_再構築_(false);
  } finally {
    cache.remove('MISAGYO_SYNCING');
    lock.releaseLock();
  }
}

// ---------- onEdit から呼ばれる ----------
// 1〜2行目(検索セル) → 自動絞り込み
// データ行のC入荷日/D入荷数 → 発注リスト大邱データと同じ自動動作(入荷日補完・連動クリア・Aチェック)
//                             ＋発注NO照合で発注リスト大邱データへ即時書き戻し＋行を黄色に
function 大邱未作業_onEdit_(e) {
  const cfg = MISAGYO_CFG;
  const range = e.range;

  // --- 検索セル(1〜2行目)の編集 → 自動絞り込み ---
  if (range.getRow() <= cfg.DST_FILTER_ROW) {
    const keywordCol = 2;
    const touchesKeyword =
      range.getRow() <= cfg.DST_KEYWORD_ROW &&
      range.getLastRow() >= cfg.DST_KEYWORD_ROW &&
      range.getColumn() <= keywordCol &&
      range.getLastColumn() >= keywordCol;

    // B1へ貼り付けた検索語から半角・全角スペース、タブ、改行を除去する。
    if (touchesKeyword) {
      const keywordCell = range.getSheet().getRange(cfg.DST_KEYWORD_ROW, keywordCol);
      const rawKeyword = keywordCell.getDisplayValue();
      const normalizedKeyword = 大邱未作業_検索語を正規化_(rawKeyword);
      if (rawKeyword !== normalizedKeyword) keywordCell.setValue(normalizedKeyword);
    }

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(1000)) return;
    try {
      大邱未作業_再構築_(false);
    } finally {
      lock.releaseLock();
    }
    return;
  }

  // --- データ行の C入荷日(3)/D入荷数(4)/Jオプション(10)/O weight(15) の編集だけ反応 ---
  if (range.getLastRow() < cfg.DST_DATA_START) return;
  const sc = range.getColumn(), ec = range.getLastColumn();
  const hitsQty = (sc <= 4 && 4 <= ec);
  const hitsDate = (sc <= 3 && 3 <= ec);
  const hitsOpt = (sc <= 10 && 10 <= ec);  // J オプション
  const hitsWgt = (sc <= 15 && 15 <= ec);  // O weight(g)
  if (!hitsQty && !hitsDate && !hitsOpt && !hitsWgt) return;

  const sh = range.getSheet();
  // 発注リスト大邱データと同じ共通処理: D入荷数→C入荷日を今日で補完・連動クリア・Aチェック(F発注NO必須)
  if ((hitsQty || hitsDate) && typeof _autoFillArrivalDateFromQty_ === 'function') {
    _autoFillArrivalDateFromQty_(sh, range, 4, 3, cfg.DST_DATA_START, 1, 6);
  }
  // 編集した行の「編集した項目だけ」を発注リスト大邱データへ書き戻し(発注NO照合)＋色を更新
  //（編集していない項目は書かない=古い表示で元データを上書きしない）
  const from = Math.max(cfg.DST_DATA_START, range.getRow());
  const to = Math.max(from, range.getLastRow());
  大邱未作業_入荷同期_(sh, from, to, {
    notify: false,
    fields: { cd: (hitsQty || hitsDate), opt: hitsOpt, wgt: hitsWgt },
    pushEmpty: true // セルを消したときは元データも消す(連動クリア)
  });
}

function 大邱未作業_検索語を正規化_(value) {
  return String(value == null ? '' : value).replace(/[\s\u3000]+/g, '');
}

// ============================================================
// 大邱未作業データの入力を発注リスト大邱データへ書き戻す(発注NO照合)
//   対象: C入荷日/D入荷数(セット) ・ Jオプション ・ O weight(g)
//   ・fromRow〜toRow を対象。値が変わった行だけ書く
//   ・入荷数が変わった行: あり→A列チェックON / なし→OFF(元シートのonEditと同じ意味)
//   ・行の色: 入荷数あり=黄色(処理済みの目印)。次のリスト更新で未作業からは消える
//   opts:
//     fields    = {cd, opt, wgt} 書き戻す項目(省略=全部)
//     pushEmpty = true: 空欄も書く(消す操作を反映) / false: 空欄は書かない(一括同期の安全用)
// ============================================================
function 大邱未作業_入荷同期_(view, fromRow, toRow, opts) {
  const cfg = MISAGYO_CFG;
  const o = opts || {};
  const notify = !!o.notify;
  const fields = o.fields || { cd: true, opt: true, wgt: true };
  const pushEmpty = o.pushEmpty !== false;
  const toggleCheck = o.check !== false; // false: 発注リスト大邱データのA列チェックを操作しない
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(cfg.SRC_SHEET);
  if (!src || !view) return 0;
  const noKey = v => (typeof EMS_転送購入Noキー_ === 'function') ? EMS_転送購入Noキー_(v) : String(v || '').trim();
  const codeKey = v => (typeof 大邱_表示コード_ === 'function') ? 大邱_表示コード_(v) : String(v || '').trim();
  const qtyNum = v => { const n = Number(String(v == null ? '' : v).replace(/,/g, '').trim()); return isFinite(n) ? n : 0; };
  const isEmpty = v => (v === '' || v == null);

  const last = Math.min(toRow, view.getLastRow());
  if (last < fromRow) return 0;
  const n = last - fromRow + 1;
  const vVals = view.getRange(fromRow, 1, n, 19).getValues(); // A..S

  // 行の色(処理の目印): 入荷数あり=黄 / なし=色なし
  const width = cfg.SRC_COL_END - cfg.SRC_COL_START + 1;
  const rowBg = vVals.map(r => new Array(width).fill(qtyNum(r[3]) > 0 ? '#ffff00' : null));
  view.getRange(fromRow, 2, n, width).setBackgrounds(rowBg);

  const targets = [];
  for (let i = 0; i < n; i++) {
    const no = noKey(vVals[i][5]); // F 発注NO
    if (!no) continue;
    targets.push({
      no: no, code: codeKey(vVals[i][10]),
      date: vVals[i][2], qty: vVals[i][3],   // C/D
      opt: vVals[i][9],                      // J オプション
      weight: vVals[i][14]                   // O weight(g)
    });
  }
  if (!targets.length) { if (notify) ss.toast('同期対象(発注NOあり)の行がありません。'); return 0; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { ss.toast('他の処理が実行中です。もう一度お試しください。'); return 0; }
  let updated = 0, missing = 0;
  try {
    const sVals = src.getDataRange().getValues();
    const byNo = {};
    for (let i = 0; i < sVals.length; i++) { const no = noKey(sVals[i][5]); if (no) (byNo[no] = byNo[no] || []).push(i); }

    const eq = (a, b) => {
      const ea = isEmpty(a), eb = isEmpty(b);
      if (ea || eb) return ea && eb;
      if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
      return String(a).trim() === String(b).trim();
    };

    // データ行(6行目〜)だけ読み書きする(ヘッダーや数式のあるかもしれない上部は触らない)
    const dataStart = 6, lastSrc = src.getLastRow();
    if (lastSrc >= dataStart) {
      const nSrc = lastSrc - dataStart + 1;
      const cd = fields.cd ? src.getRange(dataStart, 3, nSrc, 2).getValues() : null;   // C:D
      const opt = fields.opt ? src.getRange(dataStart, 10, nSrc, 1).getValues() : null; // J
      const wgt = fields.wgt ? src.getRange(dataStart, 15, nSrc, 1).getValues() : null; // O
      const checkOn = [], checkOff = [];
      let cdChanged = false, optChanged = false, wgtChanged = false;

      targets.forEach(t => {
        const rows = (byNo[t.no] || []).filter(ri => ri + 1 >= dataStart);
        if (!rows.length) { missing++; return; }
        let hit = rows[0];
        if (rows.length > 1) { const m = rows.filter(ri => codeKey(sVals[ri][10]) === t.code); if (m.length) hit = m[0]; }
        const ci = hit + 1 - dataStart;
        let rowChanged = false;

        if (cd && (pushEmpty || !isEmpty(t.date) || !isEmpty(t.qty))) {
          if (!eq(cd[ci][0], t.date) || !eq(cd[ci][1], t.qty)) {
            const qtyChanged = !eq(cd[ci][1], t.qty);
            cd[ci][0] = t.date; cd[ci][1] = t.qty; cdChanged = true; rowChanged = true;
            if (qtyChanged && toggleCheck) (qtyNum(t.qty) > 0 ? checkOn : checkOff).push('A' + (hit + 1));
          }
        }
        if (opt && (pushEmpty || !isEmpty(t.opt)) && !eq(opt[ci][0], t.opt)) {
          opt[ci][0] = t.opt; optChanged = true; rowChanged = true;
        }
        if (wgt && (pushEmpty || !isEmpty(t.weight)) && !eq(wgt[ci][0], t.weight)) {
          wgt[ci][0] = t.weight; wgtChanged = true; rowChanged = true;
        }
        if (rowChanged) updated++;
      });

      if (cdChanged) src.getRange(dataStart, 3, nSrc, 2).setValues(cd);
      if (optChanged) src.getRange(dataStart, 10, nSrc, 1).setValues(opt);
      if (wgtChanged) src.getRange(dataStart, 15, nSrc, 1).setValues(wgt);
      if (checkOn.length) src.getRangeList(checkOn).setValue(true);
      if (checkOff.length) src.getRangeList(checkOff).setValue(false);
      // 入荷日/入荷数を書き戻した → 送信済みのEMS大邱行(A列空欄)へも入荷日を反映
      if (cdChanged && typeof EMS大邱_入荷日補完_ === 'function') EMS大邱_入荷日補完_();
    }
  } finally {
    lock.releaseLock();
  }
  if (notify || updated || missing) {
    ss.toast('発注リスト大邱データへ同期: ' + updated + '件' + (missing ? ' / 発注NO不一致 ' + missing + '件' : ''), '📥 入荷同期', 4);
  }
  return updated;
}

// ============================================================
// ボタン用: チェック行の入荷数(D)を数量(L)にして入荷日(C)を今日で補完
//（発注リスト大邱データの「チェック行:入荷数を発注数量にする」と同じ機能の未作業データ版。
//   処理後はチェックを外し、発注リスト大邱データへも自動で書き戻す）
// ============================================================
function 大邱未作業_チェック行の入荷数を発注数量にする() {
  const cfg = MISAGYO_CFG;
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const view = ss.getSheetByName(cfg.DST_SHEET);
  if (!view) { ui.alert('「' + cfg.DST_SHEET + '」がありません。先に「未作業データを更新」を実行してください。'); return; }

  const start = cfg.DST_DATA_START, lastRow = view.getLastRow();
  if (lastRow < start) { ss.toast('データがありません。'); return; }
  const n = lastRow - start + 1;

  const checkR = view.getRange(start, 1, n, 1), checks = checkR.getValues();  // A チェック
  const dateR = view.getRange(start, 3, n, 1), dates = dateR.getValues();     // C 入荷日
  const qtyR = view.getRange(start, 4, n, 1), qtys = qtyR.getValues();        // D 入荷数
  const orders = view.getRange(start, 12, n, 1).getValues();                  // L 数量
  const today = (typeof _todayOnly_ === 'function')
    ? _todayOnly_()
    : (() => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; })();

  let checked = 0, updated = 0, skippedNoQty = 0;
  for (let i = 0; i < n; i++) {
    if (checks[i][0] !== true) continue;
    checked++;
    const oq = orders[i][0];
    if (oq === '' || oq == null) { skippedNoQty++; continue; }
    qtys[i][0] = oq;                                        // 入荷数 = 数量
    if (dates[i][0] === '' || dates[i][0] == null) dates[i][0] = today; // 入荷日が空なら今日
    updated++;
    // ※チェックは外さない: そのまま「チェックしたデータをEMS大邱作業に送信」を押して送れるように
  }
  if (!checked) { ss.toast('チェックされた行がありません。'); return; }

  qtyR.setValues(qtys);
  dateR.setValues(dates);

  // 発注リスト大邱データへ書き戻し(入荷日・入荷数のみ・空欄は書かない・チェックは触らない)
  大邱未作業_入荷同期_(view, start, lastRow, { notify: false, fields: { cd: true }, pushEmpty: false, check: false });

  ss.toast('入荷数を数量にしました ' + updated + '行 / 数量なしスキップ ' + skippedNoQty +
    '行（チェックは残っています。続けて送信ボタンで送れます）', '📦 入荷数=数量', 6);
}

// ボタン/メニュー用: 大邱未作業データ全行の 入荷日・入荷数・オプション・weight を発注リスト大邱データへ同期
// ※空欄は書き戻さない(古い表示や未入力で元データを消さないための安全仕様)。
//   消したいときはそのセルを編集(Delete)すれば、その行だけは即時同期で消える。
function 大邱未作業_入荷を発注リスト大邱へ同期() {
  const cfg = MISAGYO_CFG;
  const ss = SpreadsheetApp.getActive();
  const view = ss.getSheetByName(cfg.DST_SHEET);
  if (!view) { SpreadsheetApp.getUi().alert('「' + cfg.DST_SHEET + '」がありません。'); return; }
  大邱未作業_入荷同期_(view, cfg.DST_DATA_START, view.getLastRow(), { notify: true, pushEmpty: false });
}

// ---------- 数値パース（カンマ・空白対応） ----------
function 大邱未作業_数値_(v) {
  const s = String(v == null ? '' : v).replace(/,/g, '').trim();
  if (!s) return 0;
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

// ---------- 本体 ----------
// clearFilters=true: 検索条件を無視して全件表示
// 返り値: { total: 未作業件数, shown: 表示件数 }
function 大邱未作業_再構築_(clearFilters) {
  const cfg = MISAGYO_CFG;
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(cfg.SRC_SHEET);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + cfg.SRC_SHEET + '」が見つかりません。');
    return null;
  }

  const width = cfg.SRC_COL_END - cfg.SRC_COL_START + 1; // 18列 (B..S)

  // --- ソース読み込み ---
  const lastRow = src.getLastRow();
  let values = [], display = [], formats = [], backgrounds = [], soshin = [];
  if (lastRow >= cfg.SRC_DATA_START) {
    const numRows = lastRow - cfg.SRC_DATA_START + 1;
    const dataRange = src.getRange(cfg.SRC_DATA_START, cfg.SRC_COL_START, numRows, width);
    values = dataRange.getValues();
    display = dataRange.getDisplayValues();
    formats = dataRange.getNumberFormats(); // 表示形式（日付・カンマ等）も元シートからそのまま持ってくる
    backgrounds = src
      .getRange(cfg.SRC_DATA_START, cfg.COLOR_CHECK_COL, numRows, 1)
      .getBackgrounds();
    soshin = src
      .getRange(cfg.SRC_DATA_START, cfg.SRC_COL_SOSHIN, numRows, 1)
      .getDisplayValues();
  }

  // --- 未作業（EMS未積載）だけ抽出 ---
  // 幅18(B..S)内のindex: D入荷数=2, F発注NO=4, I商品名=7, K商品コード=9, L数量=10
  const requiredIdx = [4, 7, 9];
  const unworked = [];
  for (let i = 0; i < values.length; i++) {
    // ① 手塗りの背景色（オレンジ/黄色など）が付いた行は積載済み
    const bg = String(backgrounds[i][0] || '#ffffff').toLowerCase();
    if (bg !== '#ffffff' && bg !== 'white' && bg !== '') continue;
    // ② 入荷済み(D入荷数>0 × L数量>0)の行の扱い:
    //    ・EMS大邱作業データへ送信済み（Y列 EMS送信済 > 0）→ リストから外す
    //    ・まだ送っていない（Yが空/0）→ 黄色のままリストに残す
    //      ＝「入荷数入力→重量入力→チェック→送信」が終わるまで消えない
    const numD = 大邱未作業_数値_(display[i][2]);
    const numL = 大邱未作業_数値_(display[i][10]);
    if (numD > 0 && numL > 0) {
      const ySent = 大邱未作業_数値_(soshin[i] ? soshin[i][0] : '');
      if (ySent > 0) continue; // EMS大邱へ送った行（全量・一部問わず）は未作業から外す
    }
    // ③ 発注NO/商品名/商品コードが全部空の行は対象外
    if (!requiredIdx.some(idx => String(display[i][idx] || '').trim() !== '')) continue;
    unworked.push({ values: values[i], display: display[i], formats: formats[i] });
  }

  // --- 出力先シート準備 ---
  let dst = ss.getSheetByName(cfg.DST_SHEET);
  if (!dst) {
    dst = ss.insertSheet(cfg.DST_SHEET);
  }

  // --- 検索条件の取得（clearFilters時は空扱い） ---
  let keyword = '';
  let colFilters = new Array(width).fill('');
  if (!clearFilters) {
    keyword = 大邱未作業_検索語を正規化_(
      dst.getRange(cfg.DST_KEYWORD_ROW, 2).getDisplayValue()
    );
    colFilters = dst
      .getRange(cfg.DST_FILTER_ROW, 2, 1, width)
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());
  }

  const keywordTerms = keyword ? keyword.toLowerCase().split(/[\s　]+/).filter(Boolean) : [];
  const filterTerms = colFilters.map(f =>
    f ? f.toLowerCase().split(/[\s　]+/).filter(Boolean) : []
  );

  // --- 絞り込み ---
  const shownRows = unworked.filter(row => {
    const lowerCells = row.display.map(v => String(v || '').toLowerCase());

    // 全列キーワード: 各語がどこかの列に含まれる（AND）
    for (const term of keywordTerms) {
      if (!lowerCells.some(cell => cell.indexOf(term) >= 0)) return false;
    }
    // 列別: その列に全語が含まれる（AND）
    for (let c = 0; c < width; c++) {
      for (const term of filterTerms[c]) {
        if (lowerCells[c].indexOf(term) < 0) return false;
      }
    }
    return true;
  });

  // --- ヘッダー・検索エリアを整備 ---
  // 高速化: レイアウトが既にできていれば飛ばす（メニューの「更新」時は毎回整え直す）
  const needsLayout = clearFilters ||
    String(dst.getRange(1, 1).getDisplayValue()) !== '🔍' ||
    String(dst.getRange(cfg.DST_HEADER_JP, 2).getDisplayValue()).trim() === '';
  if (needsLayout) {
    const jpHeader = src.getRange(cfg.SRC_HEADER_JP, cfg.SRC_COL_START, 1, width).getDisplayValues();
    const enHeader = src.getRange(cfg.SRC_HEADER_EN, cfg.SRC_COL_START, 1, width).getDisplayValues();

    dst.getRange(1, 1).setValue('🔍').setFontWeight('bold').setFontSize(cfg.FONT_SIZE);
    dst.getRange(cfg.DST_KEYWORD_ROW, 2).setBackground('#fff2cc').setFontSize(cfg.FONT_SIZE); // B1 黄色（全列検索）
    dst.getRange(1, 3).setValue(
      '←B1は全列検索（空白・改行を除去して部分一致）／下の水色行は列ごとの条件。セルを編集すると自動で絞り込み。A列にチェックしてメニュー「大邱未作業」→「チェック行をEMS大邱へ送る」'
    ).setFontColor('#888888').setFontSize(cfg.FONT_SIZE); // 1行目も13ptにそろえる

    dst.getRange(cfg.DST_FILTER_ROW, 2, 1, width).setBackground('#e8f0fe').setFontSize(cfg.FONT_SIZE); // 水色 B2..S2
    dst.getRange(cfg.DST_HEADER_JP, 1).setValue('送る').setFontWeight('bold').setBackground('#efefef').setFontSize(cfg.FONT_SIZE);
    dst.getRange(cfg.DST_HEADER_EN, 1).setValue('✓').setFontWeight('bold').setBackground('#efefef').setFontSize(cfg.FONT_SIZE);
    dst.getRange(cfg.DST_HEADER_JP, 2, 1, width)
      .setValues(jpHeader)
      .setFontWeight('bold')
      .setBackground('#efefef')
      .setFontSize(cfg.FONT_SIZE);
    dst.getRange(cfg.DST_HEADER_EN, 2, 1, width)
      .setValues(enHeader)
      .setFontWeight('bold')
      .setBackground('#efefef')
      .setFontSize(cfg.FONT_SIZE);
    dst.setFrozenRows(cfg.DST_HEADER_EN);

    // 列幅: A=チェック用に狭く、B..Sはソースに合わせる（=同じ列アルファベット）
    dst.setColumnWidth(1, 40);
    for (let c = 0; c < width; c++) {
      dst.setColumnWidth(c + 2, src.getColumnWidth(cfg.SRC_COL_START + c));
    }
  }

  // --- データ書き込み（値のみ＝色なし） ---
  const maxRows = dst.getMaxRows();
  if (maxRows >= cfg.DST_DATA_START) {
    const dataArea = dst.getRange(cfg.DST_DATA_START, 1, maxRows - cfg.DST_DATA_START + 1, dst.getMaxColumns());
    dataArea.clearContent();
    // 書式もクリアする。旧レイアウトの日付書式が残っていると、
    // 数値の発注日(20260605)が「57371/08/25」のような異常な日付表示になる。
    dataArea.clearFormat();
  }
  if (shownRows.length) {
    const body = dst.getRange(cfg.DST_DATA_START, 2, shownRows.length, width);
    body.setValues(shownRows.map(r => r.values));
    body.setNumberFormats(shownRows.map(r => r.formats)); // 表示形式は元シートと同じにする
    body.setFontSize(cfg.FONT_SIZE).setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // 文字13pt・縦中央・折返しなしで行高を一定に
    // 入荷数が入っている行（入荷済み・未送信）は黄色で表示（作業中の目印を再構築後も維持）
    body.setBackgrounds(shownRows.map(r =>
      new Array(width).fill(大邱未作業_数値_(r.display[2]) > 0 ? '#ffff00' : null)));
    // A列チェックボックス（データ行だけ・毎回未チェックで開始）
    dst.getRange(cfg.DST_DATA_START, 1, shownRows.length, 1).insertCheckboxes();
    dst.setRowHeights(cfg.DST_DATA_START, shownRows.length, cfg.ROW_HEIGHT);

    // カート(発注NOの枝番を除いた単位 例: 20260622_13_1→20260622_13)ごとに太い下線で区切る
    const lastColLetter = String.fromCharCode(65 + width); // A + 18列 = S
    dst.getRange(cfg.DST_DATA_START, 1, shownRows.length, width + 1)
      .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID); // 全体に薄い格子
    const cartOf = v => {
      if (typeof 大邱発注_カートグループキー_ === 'function') {
        const k = 大邱発注_カートグループキー_(v);
        if (k) return k;
      }
      const s = String(v == null ? '' : v).trim();
      const m = s.match(/^(.+)_\d+$/);
      return m ? m[1] : s;
    };
    const carts = shownRows.map(r => cartOf(r.values[4])); // F列(発注NO) = B..S内のindex4
    const cartEnds = [];
    for (let i = 0; i < carts.length; i++) {
      if (i === carts.length - 1 || carts[i + 1] !== carts[i]) {
        const rowNo = cfg.DST_DATA_START + i;
        cartEnds.push('A' + rowNo + ':' + lastColLetter + rowNo);
      }
    }
    if (cartEnds.length) {
      dst.getRangeList(cartEnds).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }
  // データより下に残った古いチェックボックスの検証を外す
  const below = maxRows - (cfg.DST_DATA_START + shownRows.length) + 1;
  if (below > 0) {
    dst.getRange(cfg.DST_DATA_START + shownRows.length, 1, below, 1).clearDataValidations();
  }

  // --- 件数表示 ---
  dst.getRange(1, width + 1).clearContent(); // 旧バージョンが書いた件数表示が残っていたら消す
  dst.getRange(1, width + 2)
    .setValue('該当 ' + shownRows.length + ' / 未作業 ' + unworked.length + '件')
    .setFontColor('#666666')
    .setFontWeight('bold')
    .setFontSize(cfg.FONT_SIZE);

  SpreadsheetApp.flush();
  大邱未作業_同期解除_(); // 最新化できたので「要同期」フラグを下ろす
  return { total: unworked.length, shown: shownRows.length };
}

// ============================================================
// 大邱未作業データのチェック行を EMS大邱作業データへ送る
//   ・行の特定は発注NO（F列・一意採番済み）で発注リスト大邱データと照合
//   ・送信処理・検証は 大邱発注_チェック行をEMS大邱へ送る と同じ共通処理
//     （大邱_EMS大邱へ追記_）を使う
//   ・送信後は残り数量を再計算してリストを更新（積載済みになった行は消える）
// ============================================================
function 大邱未作業_チェック行をEMS大邱へ送る() {
  const cfg = MISAGYO_CFG;
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const view = ss.getSheetByName(cfg.DST_SHEET);
  const src = ss.getSheetByName(cfg.SRC_SHEET);
  const ems = ss.getSheetByName(
    (typeof DAEGU_CFG !== 'undefined' && DAEGU_CFG.EMS_SRC) ? DAEGU_CFG.EMS_SRC : 'EMS大邱作業データ'
  );
  if (!view) { ui.alert('「' + cfg.DST_SHEET + '」がありません。先に「未作業データを更新」を実行してください。'); return; }
  if (!src || !ems) { ui.alert('シートが見つかりません。'); return; }
  if (typeof 大邱_EMS大邱へ追記_ !== 'function' || typeof EMS_転送購入Noキー_ !== 'function') {
    ui.alert('送信用の共通処理が見つかりません（【大邱】データ転送.gs を確認してください）。');
    return;
  }

  const t0 = new Date();
  const L = msg => Logger.log('[未作業→EMS大邱] ' + msg);
  L('開始 ' + Utilities.formatDate(t0, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));

  // --- ① 未作業シートのチェック行を収集（ロック不要の読み取り） ---
  const lastRow = view.getLastRow();
  if (lastRow < cfg.DST_DATA_START) { ss.toast('送る行がありません。'); return; }
  const vVals = view.getRange(cfg.DST_DATA_START, 1, lastRow - cfg.DST_DATA_START + 1, 19).getValues(); // A..S
  const checked = [];
  for (let i = 0; i < vVals.length; i++) {
    if (vVals[i][0] !== true) continue; // A列チェック
    const no = EMS_転送購入Noキー_(vVals[i][5]);                 // F 発注NO
    const code = (typeof 大邱_表示コード_ === 'function') ? 大邱_表示コード_(vVals[i][10]) : String(vVals[i][10] || '').trim(); // K
    checked.push({ viewRow: cfg.DST_DATA_START + i, no: no, code: code });
  }
  if (checked.length === 0) { ss.toast('チェックされた行がありません。'); return; }
  L('チェック行: ' + checked.length + '件');

  // --- ② 発注リスト大邱データを発注NOで照合して送信対象を作る（元データの最新値を使う） ---
  const sVals = src.getDataRange().getValues();
  const byNo = {};
  for (let i = 0; i < sVals.length; i++) {
    const no = EMS_転送購入Noキー_(sVals[i][5]); // F
    if (!no) continue;
    (byNo[no] = byNo[no] || []).push(i);
  }

  const picked = [];
  const missing = [];
  for (const c of checked) {
    if (!c.no) { missing.push('行' + c.viewRow + ': 発注NOが空'); continue; }
    const rows = byNo[c.no] || [];
    let hit = -1;
    if (rows.length === 1) {
      hit = rows[0];
    } else if (rows.length > 1) {
      // 同じ発注NOが複数ある場合は商品コードでも照合
      const byCode = rows.filter(ri => {
        const sc = (typeof 大邱_表示コード_ === 'function') ? 大邱_表示コード_(sVals[ri][10]) : String(sVals[ri][10] || '').trim();
        return sc === c.code;
      });
      hit = byCode.length ? byCode[0] : rows[0];
      L('⚠ 発注NO重複 ' + c.no + ' → 行' + (hit + 1) + 'を採用');
    }
    if (hit < 0) { missing.push(c.no + '（発注リスト大邱データに見つからない）'); continue; }

    const r = sVals[hit];
    const no = EMS_転送購入Noキー_(r[5]);
    const code = (typeof 大邱_表示コード_ === 'function') ? 大邱_表示コード_(r[10]) : String(r[10] || '').trim();
    if (!no || !code) { missing.push(c.no + '（購入No/商品コード空）'); continue; }
    picked.push({
      row: hit + 1, no: no, code: code,
      qty: (typeof EMS_表示数量_ === 'function') ? EMS_表示数量_(r[11]) : r[11],
      date: r[2], vendor: r[7], name: r[8], item: r[13], weight: r[14], price: r[15],
      viewRow: c.viewRow
    });
  }
  L('送信対象: ' + picked.length + '件 / 照合不可: ' + missing.length + '件');
  missing.forEach(m => L('  ✗ ' + m));

  if (picked.length === 0) {
    ui.alert('送信できる行がありません。\n\n' + missing.slice(0, 10).join('\n'));
    return;
  }

  // --- ③ 確認ダイアログ（ロック無し） ---
  const preview = picked.slice(0, 12).map(p => `${p.no} / ${p.code} / ${p.qty || 0}個`).join('\n');
  const res = ui.alert('未作業リストからEMS大邱へ送る',
    `${picked.length}件をEMS大邱作業データの最終行の下へ送ります。\n\n${preview}` +
    (picked.length > 12 ? '\n…ほか' : '') +
    (missing.length ? '\n\n⚠ 照合できず送らない行: ' + missing.length + '件' : '') +
    '\n\n（EMS番号・発送日はEMS担当が記入）\n実行する？',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) { L('終了: ユーザーがキャンセル'); ui.alert('やめました。'); return; }

  // --- ④ ロック取得（最大30秒待つ）→ 書き込み ---
  const lock = (typeof 大邱_ロック取得_ === 'function')
    ? 大邱_ロック取得_(30000, '未作業→EMS大邱')
    : LockService.getDocumentLock();
  if (!lock) { L('中断: ロック取得タイムアウト'); return; }
  try {
    const result = 大邱_EMS大邱へ追記_(ems, picked, L);
    const n = result.n;

    // 未作業シートのチェックを外す（1回のAPI呼び出し）
    view.getRangeList(picked.map(p => 'A' + p.viewRow)).setValue(false);

    // 残り数量を再計算してから未作業リストを最新化（積載済みになった行は自動で消える）
    if (typeof 大邱発注_チェックと残り数量を設置 === 'function') {
      大邱発注_チェックと残り数量を設置();
    }
    大邱未作業_再構築_(false);

    const allOK = (result.received === n && result.ngCount === 0);
    const mark = allOK ? '✅ 件数一致しました。' : '⚠️ 件数不一致あり！実行ログを確認してください。';
    L('完了: 送った ' + n + '件 / 送られた ' + result.received + '件 / 所要 ' + (new Date() - t0) + 'ms');
    ui.alert('EMS大邱へ送信 完了',
      '送った件数　: ' + n + '件\n' +
      '送られた件数: ' + result.received + '件\n' +
      '照合　　　　: 件数一致 ' + result.okCount + ' / 件数不一致 ' + result.ngCount +
      (missing.length ? '\n照合不可で未送信: ' + missing.length + '件' : '') +
      '\n\n' + mark,
      ui.ButtonSet.OK);
  } catch (err) {
    L('🛑 例外: ' + (err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}
