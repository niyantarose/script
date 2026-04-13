/******************************************************
 * ✅ チェックボックス操作パネル（U列）→ 予約だけ書く（超安全版）
 * - onEdit(インストール型) で即座に「ジョブをキューに追加」
 * - 実行は「時間主導トリガー（毎分）」のランナーが担当
 * - キュー方式なので、連打/複数ジョブでも消えない
 * - 全JOB完了/エラー時に通知（toast）
 ******************************************************/

const PANEL_SHEET_NAME = '①商品入力シート';
const PANEL_CHECK_COL  = 21; // U列

// 行番号定義（チェックボックスがある行）
const ROW_1  = 3;
const ROW_2  = 4;
const ROW_3  = 5;
const ROW_4  = 6;
const ROW_5  = 7;
const ROW_6  = 8;
const ROW_7  = 9;
const ROW_8  = 10;
const ROW_9  = 11;
const ROW_10 = 12;
const ROW_11 = 13;

// ⑧⑨の対象シート名（必要ならそのまま）
const SHEET_FOR_8 = '②オプション入力シート';
const SHEET_FOR_9 = 'Yahoo商品登録シート';

// キュー用プロパティキー
const CB_QUEUE_KEY = 'CB_JOB_QUEUE_JSON';

// ロック
const USE_LOCK = true;
const LOCK_WAIT_MS = 3000;

/**
 * ✅ インストール型トリガーでこの関数を「編集時」に指定する
 * （時計アイコン → トリガー追加 → 関数=CB_onEdit）
 */
function CB_onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== PANEL_SHEET_NAME) return;
  if (e.range.getColumn() !== PANEL_CHECK_COL) return;

  // チェックボックスは e.value が 'TRUE' / 'FALSE'
  if (e.value !== 'TRUE') return;

  let lock;
  if (USE_LOCK) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_WAIT_MS)) {
      // 反応がないと不安なので、押した瞬間OFFに戻す
      e.range.setValue(false);
      try { e.source.toast('連打防止：少し待ってください', '受付待ち', 3); } catch (_) {}
      return;
    }
  }

  try {
    // 押したら即OFF（ボタン化）
    e.range.setValue(false);

    const row = e.range.getRow();
    const job = rowToJob_(row);

    const panelSheet = e.source.getSheetByName(PANEL_SHEET_NAME);
    if (!job) {
      logToX11_(panelSheet, `[受付NG] この行(${row})は未定義`);
      try { e.source.toast(`この行（${row}）には機能がありません`, 'エラー', 4); } catch (_) {}
      return;
    }

    // キューに追加
    enqueueJob_(job);

    const label = jobLabel_(job);
    logToX11_(panelSheet, `[予約] ${label} をキューに追加`);
    try { e.source.toast(`予約OK: ${label}`, '受付OK', 4); } catch (_) {}

  } catch (err) {
    console.error(err);
    try { e.source.toast(`エラー: ${err}`, 'システムエラー', 8); } catch (_) {}
    try {
      const ss = e.source;
      const panelSheet = ss.getSheetByName(PANEL_SHEET_NAME);
      logToX11_(panelSheet, `[予約エラー] ${err && err.message ? err.message : err}`);
    } catch (_) {}
  } finally {
    if (lock) try { lock.releaseLock(); } catch (_) {}
  }
}

/** ✅毎分起動するランナー（時間主導トリガーで動かす） */
function CB_実行ランナー_毎分() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) return; // 二重起動を止める

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const panelSheet = ss.getSheetByName(PANEL_SHEET_NAME);
    if (!panelSheet) return;

    // 生存確認（任意）
    panelSheet.getRange('V2').setValue(new Date());

    // 1件取り出し（FIFO）
    const job = dequeueJob_();
    if (!job) return;

    logToX11_(panelSheet, `[実行] START: ${jobLabel_(job)}`);

    try {
      runJob_(job);

      logToX11_(panelSheet, `[完了] SUCCESS: ${jobLabel_(job)}`);
      notifyDone_(ss, panelSheet, job, true);

    } catch (err) {
      console.error(err);
      logToX11_(panelSheet, `[エラー] ERROR: ${jobLabel_(job)} / ${err && err.message ? err.message : err}`);
      notifyDone_(ss, panelSheet, job, false, err);
    }

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * 完了/エラー通知（toast + ログ）
 */
function notifyDone_(ss, panelSheet, job, ok, err) {
  try {
    const label = jobLabel_(job);
    const title = ok ? '✅ 完了' : '❌ エラー';
    const msg = ok
      ? `完了しました: ${label}`
      : `失敗: ${label}\n${(err && err.message) ? err.message : err}`;

    // 画面右下のトースト
    ss.toast(msg, title, ok ? 5 : 10);

    // ログにも残す
    logToX11_(panelSheet, `[通知] ${title}: ${label}`);

  } catch (_) {
    // toastできない状況でも落とさない
  }
}

/**
 * JOB名 → 日本語ラベル変換
 */
function jobLabel_(job) {
  const map = {
    JOB_1:  '①配送テキスト反映',
    JOB_2:  '②全商品リネーム',
    JOB_3:  '③商品説明整形チェック',
    JOB_4:  '④画像URL挿入(ST)',
    JOB_5:  '⑤ST画像URLサーバー変換',
    JOB_6:  '⑥プロダクトカテゴリ入力',
    JOB_7:  '⑦サブコード&オプション生成',
    JOB_8:  '⑧Yahoo反映→CSV成形',
    JOB_9:  '⑨Yahoo商品登録シートCSVダウンロード',
    JOB_10: '⑩YahooUpload一括',
    JOB_11: '⑪Yahoo再アップロード一括'
  };
  return map[job] || job;
}

// ★⑧専用ラッパー（名前が絶対に被らない）
function JOB8_YahooCSV登録用に型を整える__RUN() {
  商品入力からYahooに反映してCSV成形_一括();
}

/** ジョブ実行本体 */
function runJob_(job) {
  switch (job) {
    case 'JOB_1':  配送テキストを反映_商品入力シート対象(); break;
    case 'JOB_2':  全商品一括リネーム(); break;
    case 'JOB_3':  商品説明を一括整形してチェック(); break;

    case 'JOB_4':  商品画像URLを挿入_ST_商品入力シート対象(); break;

    case 'JOB_5':
      ST列画像URLをサーバー変換_汎用_バッチ_6分制限対策版_(
        '①商品入力シート',
        2,
        true
      );
      break;

    case 'JOB_6':  プロダクトカテゴリを入力(); break;
    case 'JOB_7':  サブコードとオプションを生成(); break;

    case 'JOB_8':
      runOnSheet_(SHEET_FOR_8, () => JOB8_YahooCSV登録用に型を整える__RUN());
      break;

    case 'JOB_9':
      runOnSheet_(SHEET_FOR_9, () => Yahoo商品登録シートCSVダウンロード());
      break;

    case 'JOB_10': YahooUpload_一括実行('AUTO'); break;
    case 'JOB_11': Yahoo再アップロード_一括実行(); break;

    default:
      throw new Error('未定義のジョブ: ' + job);
  }
}

/* =========================
 * キュー（FIFO）実装
 * ========================= */

function enqueueJob_(job) {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty(CB_QUEUE_KEY);
  let q = [];
  if (json) {
    try { q = JSON.parse(json) || []; } catch (_) { q = []; }
  }

  // 同じジョブがすでに末尾にいたら重複追加しない（好みで外してOK）
  if (q.length > 0 && q[q.length - 1] === job) {
    return;
  }

  q.push(job);

  // キュー肥大防止（上限100）
  if (q.length > 100) q = q.slice(q.length - 100);

  props.setProperty(CB_QUEUE_KEY, JSON.stringify(q));
}

function dequeueJob_() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty(CB_QUEUE_KEY);
  if (!json) return null;

  let q = [];
  try { q = JSON.parse(json) || []; } catch (_) { q = []; }
  if (q.length === 0) {
    props.deleteProperty(CB_QUEUE_KEY);
    return null;
  }

  const job = q.shift();
  if (q.length === 0) props.deleteProperty(CB_QUEUE_KEY);
  else props.setProperty(CB_QUEUE_KEY, JSON.stringify(q));

  return job;
}

/* =========================
 * ヘルパー
 * ========================= */

function rowToJob_(row) {
  switch (row) {
    case ROW_1:  return 'JOB_1';
    case ROW_2:  return 'JOB_2';
    case ROW_3:  return 'JOB_3';
    case ROW_4:  return 'JOB_4';
    case ROW_5:  return 'JOB_5';
    case ROW_6:  return 'JOB_6';
    case ROW_7:  return 'JOB_7';
    case ROW_8:  return 'JOB_8';
    case ROW_9:  return 'JOB_9';
    case ROW_10: return 'JOB_10';
    case ROW_11: return 'JOB_11';
    default:     return null;
  }
}

function runOnSheet_(sheetName, fn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = ss.getActiveSheet();
  const target = ss.getSheetByName(sheetName);
  if (!target) throw new Error('シートが見つかりません: ' + sheetName);

  ss.setActiveSheet(target);
  try { fn(); } finally { ss.setActiveSheet(current); }
}

/** X11セルにログ追記（パネルシート固定で使う前提） */
function logToX11_(sheet, message) {
  if (!sheet) return;
  try {
    const cell = sheet.getRange('X11');
    const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm:ss');
    const newLine = `${ts} ${message}`;

    let current = String(cell.getDisplayValue() || '');
    if (current.length > 5000) current = current.slice(0, 5000);

    cell.setValue(newLine + '\n' + current);
  } catch (e) {
    console.error('ログ書き込み失敗', e);
  }
}
