/******************************************************
 * ST列画像URLをサーバー変換（6分制限対策・高速化版）
 * v5.1: UIなし文脈でも落ちない修正版
 ******************************************************/

'use strict';

/** =====================================================
 * ★ 環境設定（今のサーバ形）
 * ===================================================== */
const IMGCV_VPS_BATCH_URL       = 'https://img.niyantarose.com/imgcv_batch.php';
const IMGCV_SERVER_URL_BASE     = 'https://img.niyantarose.com/img/';

const IMGCV_SHEET_DEFAULT  = '①商品入力シート';
const IMGCV_HEADER_DEFAULT = 2;

const IMGCV_COL_CODE = 3;
const IMGCV_COL_S    = 19;
const IMGCV_COL_T    = 20;

/** =====================================================
 * 実行制御
 * ===================================================== */
const IMGCV_MAX_ROWS_PER_RUN   = 200;
const IMGCV_BATCH_SIZE         = 30;
const IMGCV_MAX_RUN_MS         = 5 * 60 * 1000;
const IMGCV_SAFE_MARGIN_MS     = 30 * 1000;

const IMGCV_CURSOR_PREFIX = 'IMGCV_CURSOR_';

/** =====================================================
 * ログ/リトライ設定
 * ===================================================== */
const IMGCV_RETRY_MAX_PER_RUN = 1;
const IMGCV_LOG_SKIP          = true;
const IMGCV_LOG_ONLY_SUMMARY  = true;

/** =====================================================
 * VPS secret
 * ===================================================== */
const IMGCV_VPS_SECRET_PROP_KEY = 'VPS_SC';

/* =====================================================
 * 安全通知
 * ===================================================== */
function IMGCV_safeNotify_(message, title) {
  const ttl = title || '画像変換④';

  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(ttl, message, ui.ButtonSet.OK);
    return;
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      ss.toast(message, ttl, 5);
      return;
    }
  } catch (e) {}

  Logger.log(`${ttl}: ${message}`);
}

/* =====================================================
 * URL正規化（Drive）
 * ===================================================== */
function IMGCV_normalizeDriveUrl_(u) {
  u = String(u || '').trim();
  if (!u) return u;

  if (!/drive\.google\.com|drive\.usercontent\.google\.com/i.test(u)) return u;

  let m = u.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=download&id=${m[1]}`;

  m = u.match(/\/open\?id=([a-zA-Z0-9_-]+)/i);
  if (m) return `https://drive.google.com/uc?export=download&id=${m[1]}`;

  m = u.match(/[?&]id=([^&]+)/i);
  if (m) {
    const idRaw = decodeURIComponent(m[1]);

    let m2 = idRaw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (m2) return `https://drive.google.com/uc?export=download&id=${m2[1]}`;

    m2 = idRaw.match(/\/open\?id=([a-zA-Z0-9_-]+)/i);
    if (m2) return `https://drive.google.com/uc?export=download&id=${m2[1]}`;

    if (/^[a-zA-Z0-9_-]{10,}$/.test(idRaw)) {
      return `https://drive.google.com/uc?export=download&id=${idRaw}`;
    }
    return u;
  }

  if (/drive\.google\.com\/uc\?/i.test(u)) {
    u = u.replace(/export=view/ig, 'export=download');
    if (!/[?&]export=download/i.test(u)) {
      u = u.replace(/\/uc\?/i, '/uc?export=download&');
    }
    return u;
  }

  return u;
}

/* =====================================================
 * URL抽出
 * ===================================================== */
function IMGCV_pickHttpUrl_(s) {
  const m = String(s || '').match(/https?:\/\/[^\s;]+/i);
  return m ? m[0] : '';
}

/* =====================================================
 * 判定
 * ===================================================== */
function IMGCV_isServerUrl_(u) {
  return String(u || '').startsWith(IMGCV_SERVER_URL_BASE);
}
function IMGCV_isWebpUrl_(u) {
  return /\.webp(\?|$)/i.test(String(u || ''));
}

/* =====================================================
 * VPS secret取得
 * ===================================================== */
function IMGCV_getSecret_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(IMGCV_VPS_SECRET_PROP_KEY) || ''
  ).trim();
}

/* =====================================================
 * 入口（手動）
 * ===================================================== */
function ST列画像URLをサーバー変換_商品入力シート対象() {
  ST列画像URLをサーバー変換_汎用_バッチ_6分制限対策版_(
    IMGCV_SHEET_DEFAULT,
    IMGCV_HEADER_DEFAULT,
    false
  );
}

/* =====================================================
 * 自動継続
 * ===================================================== */
function IMGCV_自動継続実行() {
  ST列画像URLをサーバー変換_汎用_バッチ_6分制限対策版_(
    IMGCV_SHEET_DEFAULT,
    IMGCV_HEADER_DEFAULT,
    true
  );
}

/* =====================================================
 * 最初から
 * ===================================================== */
function ST列画像URLをサーバー変換_最初から実行() {
  IMGCV_cursorSet_(IMGCV_SHEET_DEFAULT, 0);
  IMGCV_deleteTrigger_();
  ST列画像URLをサーバー変換_商品入力シート対象();
}

/* =====================================================
 * 停止
 * ===================================================== */
function ST列画像URLをサーバー変換_停止() {
  IMGCV_deleteTrigger_();
  IMGCV_safeNotify_('自動継続を停止しました');
}

/* =====================================================
 * 進捗確認
 * ===================================================== */
function ST列画像URLをサーバー変換_進捗確認() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMGCV_SHEET_DEFAULT);
  if (!sheet) {
    IMGCV_safeNotify_('シートが見つかりません');
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataRows = Math.max(0, lastRow - IMGCV_HEADER_DEFAULT);
  const cursor = IMGCV_cursorGet_(IMGCV_SHEET_DEFAULT);
  const progress = dataRows > 0 ? ((cursor / dataRows) * 100).toFixed(1) : '0.0';
  const hasTrigger = IMGCV_hasTrigger_();

  IMGCV_safeNotify_(
    `【進捗状況】\n` +
    `総データ行数: ${dataRows}\n` +
    `処理済み行数: ${cursor}\n` +
    `進捗率: ${progress}%\n` +
    `自動継続: ${hasTrigger ? '実行中' : '停止中'}`
  );
}

/* =====================================================
 * メイン
 * ===================================================== */
function ST列画像URLをサーバー変換_汎用_バッチ_6分制限対策版_(シート名, ヘッダー行数, isAuto) {
  const startedAt = Date.now();
  const safeEnd = startedAt + IMGCV_MAX_RUN_MS - IMGCV_SAFE_MARGIN_MS;

  const secret = IMGCV_getSecret_();
  if (!secret) {
    if (!isAuto) IMGCV_safeNotify_(`ScriptPropertiesに ${IMGCV_VPS_SECRET_PROP_KEY} が未設定です`);
    IMGCV_deleteTrigger_();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(シート名);
  if (!sheet) {
    if (!isAuto) IMGCV_safeNotify_(`シートが見つかりません: ${シート名}`);
    IMGCV_deleteTrigger_();
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= ヘッダー行数) {
    if (!isAuto) IMGCV_safeNotify_('データ行がありません');
    IMGCV_deleteTrigger_();
    return;
  }

  const dataRows = lastRow - ヘッダー行数;
  let startPos = IMGCV_cursorGet_(シート名);

  if (startPos >= dataRows) {
    IMGCV_cursorSet_(シート名, 0);
    IMGCV_deleteTrigger_();
    if (!isAuto) IMGCV_safeNotify_('全行の処理が完了しました！\nカーソルをリセットしました。');
    return;
  }

  let totalOkCount = 0;
  const totalStats = IMGCV_statsInit_();
  let loopCount = 0;

  while (startPos < dataRows) {
    const now = Date.now();

    // もうほとんど時間が残ってなければ、ここで fallback
    if (now >= safeEnd || (safeEnd - now) < 15000) {
      Logger.log(`[IMGCV] 時間残り不足のため fallback へ。remaining=${dataRows - startPos}`);
      break;
    }

    const slice = IMGCV_runOneSlice_(sheet, シート名, ヘッダー行数, startPos, dataRows, safeEnd, secret);

    // 1行も進まなければ無限ループ防止で抜ける
    if (!slice.processedRows) {
      Logger.log('[IMGCV] processedRows=0 のため停止');
      break;
    }

    startPos += slice.processedRows;
    IMGCV_cursorSet_(シート名, startPos);

    totalOkCount += slice.okCount;
    IMGCV_statsMerge_(totalStats, slice.stats);
    loopCount++;

    Logger.log(`[IMGCV] loop=${loopCount} cursor=${startPos}/${dataRows}`);

    // 完了したら抜ける
    if (startPos >= dataRows) break;
  }

  const elapsed = ((Date.now() - startedAt) / 1000).toFixed(1);
  const remaining = dataRows - startPos;

  Logger.log(`[IMGCV][TOTAL SUMMARY] loops=${loopCount} OK=${totalStats.ok} FAIL=${totalStats.fail} SKIP=${totalStats.skip} okApplied=${totalOkCount} remaining=${remaining}`);

  if (remaining > 0) {
    IMGCV_ensureTrigger_();
    if (!isAuto) {
      IMGCV_safeNotify_(
        `変換成功: ${totalOkCount}\n処理時間: ${elapsed}秒\n残り行数: ${remaining}\n\n⏳ 今回は時間不足のため、残りは自動継続します`
      );
    }
  } else {
    IMGCV_deleteTrigger_();
    IMGCV_cursorSet_(シート名, 0);
    if (!isAuto) {
      IMGCV_safeNotify_(
        `変換成功: ${totalOkCount}\n処理時間: ${elapsed}秒\n\n✅ 全行の処理が完了しました！`
      );
    }
  }
}

/* =====================================================
 * 結果反映（OK + SKIP を server jpg URLに置換）
 * ===================================================== */
function IMGCV_applyResults_(okJobs) {
  let okCount = 0;

  okJobs.forEach(job => {
    try {
      const newUrl =
        IMGCV_SERVER_URL_BASE +
        encodeURIComponent(job.商品コード) +
        '/' +
        job.保存名;

      if (job.種別 === 'S') {
        if (job.元配列 && job.元配列[job.行]) {
          job.元配列[job.行][0] = newUrl;
          okCount++;
        }
      } else {
        if (job.詳細配列 && job.元配列 && job.元配列[job.行]) {
          job.詳細配列[job.詳細index] = newUrl;
          job.元配列[job.行][0] = job.詳細配列.join(';');
          okCount++;
        }
      }
    } catch (e) {
      Logger.log(`[IMGCV] 反映エラー: ${e}`);
    }
  });

  return okCount;
}

/* =====================================================
 * VPS 実行（fetchAll / tsv_b64）
 * ===================================================== */
function IMGCV_vpsRun_withTimeLimit_(jobs, secret) {
  if (!jobs.length) {
    return { okJobs: [], failJobs: [], skipJobs: [], stats: IMGCV_statsInit_() };
  }

  const chunks = IMGCV_chunk_(jobs, IMGCV_BATCH_SIZE);

  const requests = chunks.map(batch => {
    const tsv = batch
      .map(j => `${String(j.元URL || '').replace(/[\r\n]/g, '')}\t${j.商品コード}\t${j.保存名}`)
      .join('\n') + '\n';

    const tsv_b64 = Utilities.base64EncodeWebSafe(tsv);

    return {
      url: IMGCV_VPS_BATCH_URL,
      method: 'post',
      headers: { 'X-VPS-Secret': secret },
      payload: {
        secret: secret,
        tsv_b64: tsv_b64,
        force: '0'
      },
      muteHttpExceptions: true,
      followRedirects: true
    };
  });

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    Logger.log(`[IMGCV] fetchAll失敗: ${e}`);
    return { okJobs: [], failJobs: jobs.slice(), skipJobs: [], stats: IMGCV_statsInit_() };
  }

  const jobMap = new Map();
  jobs.forEach(j => jobMap.set(`${j.商品コード}\t${j.保存名}`, j));

  const ok = [];
  const fail = [];
  const skip = [];
  const stats = IMGCV_statsInit_();

  responses.forEach((resp, i) => {
    const http = resp.getResponseCode();
    const body = String(resp.getContentText() || '');

    if (http !== 200) {
      chunks[i].forEach(j => fail.push(j));
      stats.fail += chunks[i].length;
      Logger.log(`[IMGCV] batch ${i + 1}/${chunks.length} FAIL: HTTP=${http} bodyHead="${body.slice(0, 120)}"`);
      return;
    }

    const parsed = IMGCV_parseVpsBody_(body, jobMap);
    ok.push(...parsed.ok);
    fail.push(...parsed.fail);
    skip.push(...parsed.skip);
    IMGCV_statsMerge_(stats, parsed.stats);
  });

  return { okJobs: ok, failJobs: fail, skipJobs: skip, stats };
}

/* =====================================================
 * パース＆統計
 * ===================================================== */
function IMGCV_statsInit_() {
  return { ok: 0, fail: 0, skip: 0, exists: 0, cached: 0, downloaded: 0, reasons: {} };
}
function IMGCV_statsMerge_(a, b) {
  a.ok += b.ok;
  a.fail += b.fail;
  a.skip += b.skip;
  a.exists += b.exists;
  a.cached += b.cached;
  a.downloaded += b.downloaded;

  Object.keys(b.reasons).forEach(k => {
    a.reasons[k] = (a.reasons[k] || 0) + b.reasons[k];
  });
}
function IMGCV_parseVpsBody_(body, jobMap) {
  const ok = [];
  const fail = [];
  const skip = [];
  const stats = IMGCV_statsInit_();

  const lines = String(body || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  for (const line of lines) {
    if (!line.startsWith('IMGCV_RESULT')) continue;

    const parts = line.includes('\t') ? line.split('\t') : line.split(/\s+/);
    const code = parts[1] || '';
    const name = parts[2] || '';
    const status = (parts[3] || '').toUpperCase();
    const reason = parts.slice(4).join(' ').trim();

    const key = `${code}\t${name}`;
    const job = jobMap.get(key);

    stats.reasons[`${status}:${reason || '-'}`] =
      (stats.reasons[`${status}:${reason || '-'}`] || 0) + 1;

    if (status === 'OK') {
      stats.ok++;
      if (/CACHED/i.test(reason)) stats.cached++;
      else stats.downloaded++;
      if (job) ok.push(job);
    } else if (status === 'SKIP') {
      stats.skip++;
      if (/EXISTS/i.test(reason)) stats.exists++;
      if (job) skip.push(job);
      if (IMGCV_LOG_SKIP && !IMGCV_LOG_ONLY_SUMMARY) {
        Logger.log(`[IMGCV][SKIP] ${code} ${name} ${reason}`);
      }
    } else {
      stats.fail++;
      if (job) {
        job.__imgcv_reason = reason;
        fail.push(job);
      }
      Logger.log(`[IMGCV][FAIL] ${code} ${name} ${reason}`);
    }
  }
  return { ok, fail, skip, stats };
}

/* =====================================================
 * ユーティリティ
 * ===================================================== */
function IMGCV_chunk_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}
function IMGCV_cursorKey_(sheetName) {
  return IMGCV_CURSOR_PREFIX + sheetName;
}
function IMGCV_cursorGet_(sheetName) {
  const v = PropertiesService.getScriptProperties().getProperty(IMGCV_cursorKey_(sheetName));
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
function IMGCV_cursorSet_(sheetName, row) {
  PropertiesService.getScriptProperties().setProperty(IMGCV_cursorKey_(sheetName), String(row));
}

/* =====================================================
 * トリガー
 * ===================================================== */
function IMGCV_ensureTrigger_() {
  if (IMGCV_hasTrigger_()) return;
  ScriptApp.newTrigger('IMGCV_自動継続実行').timeBased().after(1 * 60 * 1000).create();
  Logger.log('[IMGCV] トリガー作成');
}
function IMGCV_deleteTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'IMGCV_自動継続実行') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
function IMGCV_hasTrigger_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'IMGCV_自動継続実行');
}

/* =====================================================
 * メニュー
 * ===================================================== */
function IMGCV_onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖼️ 画像変換④')
    .addItem('▶️ 実行（続きから）', 'ST列画像URLをサーバー変換_商品入力シート対象')
    .addItem('🔄 最初から実行', 'ST列画像URLをサーバー変換_最初から実行')
    .addItem('⏹️ 停止', 'ST列画像URLをサーバー変換_停止')
    .addItem('📊 進捗確認', 'ST列画像URLをサーバー変換_進捗確認')
    .addToUi();
}
function onOpen(e) {
  IMGCV_onOpen();
}
function onInstall(e) {
  IMGCV_onOpen();
}

/* =====================================================
 * リンクなしプレーンテキストのRichTextValue
 * ===================================================== */
function IMGCV_buildPlainRichText_(text) {
  return SpreadsheetApp.newRichTextValue()
    .setText(String(text || ''))
    .build();
}

function IMGCV_runOneSlice_(sheet, シート名, ヘッダー行数, startPos, dataRows, safeEnd, secret) {
  const rows = Math.min(IMGCV_MAX_ROWS_PER_RUN, dataRows - startPos);
  const rowTop = ヘッダー行数 + 1 + startPos;

  const codes = sheet.getRange(rowTop, IMGCV_COL_CODE, rows, 1).getValues();
  const S = sheet.getRange(rowTop, IMGCV_COL_S, rows, 1).getValues();
  const T = sheet.getRange(rowTop, IMGCV_COL_T, rows, 1).getValues();

  let S_rich = null;
  let T_rich = null;
  try {
    S_rich = sheet.getRange(rowTop, IMGCV_COL_S, rows, 1).getRichTextValues();
    T_rich = sheet.getRange(rowTop, IMGCV_COL_T, rows, 1).getRichTextValues();
  } catch (e) {}

  const jobs = [];
  let processedRows = 0;

  Logger.log(`=== IMGCV slice sheet=${シート名} start=${startPos + 1}/${dataRows} rows=${rows} ===`);

  for (let r = 0; r < rows; r++) {
    if (Date.now() > safeEnd) {
      Logger.log(`[IMGCV] 時間切れ（ジョブ作成中）: r=${r}`);
      break;
    }

    processedRows++;

    const code = String(codes[r][0] || '').trim();
    if (!code) continue;

    // ---------- S列 ----------
    const s0 = String(S[r][0] || '').trim();
    let sUrl = IMGCV_pickHttpUrl_(s0) || s0;
    if (!sUrl && S_rich && S_rich[r] && S_rich[r][0]) {
      const linkUrl = S_rich[r][0].getLinkUrl();
      if (linkUrl) sUrl = linkUrl;
    }
    const s = IMGCV_normalizeDriveUrl_(sUrl);

    if (s) {
      const isServer = IMGCV_isServerUrl_(s);
      const isWebp = IMGCV_isWebpUrl_(s);

      if (!isServer || (isServer && isWebp)) {
        jobs.push({
          種別: 'S',
          元URL: s,
          商品コード: code,
          保存名: `${code}.jpg`,
          行: r,
          元配列: S
        });
      }
    }

    // ---------- T列 ----------
    const t0Val = String(T[r][0] || '').trim();
    if (t0Val) {
      let rawText = t0Val;
      if (!IMGCV_pickHttpUrl_(rawText) && T_rich && T_rich[r] && T_rich[r][0]) {
        const linkUrl = T_rich[r][0].getLinkUrl();
        if (linkUrl) rawText = linkUrl;
      }

      const urls = rawText
        .split(/[\n;]+/)
        .map(x => IMGCV_pickHttpUrl_(x) || x.trim())
        .filter(x => /^https?:\/\//i.test(x))
        .map(x => IMGCV_normalizeDriveUrl_(x));

      urls.forEach((u, i) => {
        const isServer = IMGCV_isServerUrl_(u);
        const isWebp = IMGCV_isWebpUrl_(u);

        if (!isServer || (isServer && isWebp)) {
          jobs.push({
            種別: 'T',
            元URL: u,
            商品コード: code,
            保存名: `${code}_${i + 1}.jpg`,
            行: r,
            詳細配列: urls,
            詳細index: i,
            元配列: T
          });
        }
      });
    }
  }

  const stats = IMGCV_statsInit_();

  // 対象URLが無いだけなら、行だけ進めて返す
  if (!jobs.length) {
    Logger.log(`[IMGCV] このスライスは対象なし rows=${processedRows}`);
    return {
      processedRows,
      okCount: 0,
      stats,
      hadJobs: false
    };
  }

  Logger.log(`[IMGCV] jobs=${jobs.length}`);

  let result = IMGCV_vpsRun_withTimeLimit_(jobs, secret);
  let okJobs = result.okJobs;
  let failJobs = result.failJobs;
  let skipJobs = result.skipJobs;
  IMGCV_statsMerge_(stats, result.stats);

  for (let retry = 0; retry < IMGCV_RETRY_MAX_PER_RUN; retry++) {
    if (!failJobs.length) break;

    const retryTargets = failJobs.filter(j => !/YIMG_PLACEHOLDER|INVALID_IMAGE|TOO_SMALL/i.test(j.__imgcv_reason || ''));
    if (retryTargets.length === 0) break;

    Logger.log(`[IMGCV] retry=${retry + 1} → 再実行 ${retryTargets.length}件`);
    const retryResult = IMGCV_vpsRun_withTimeLimit_(retryTargets, secret);

    const permanentFails = failJobs.filter(j => /YIMG_PLACEHOLDER|INVALID_IMAGE|TOO_SMALL/i.test(j.__imgcv_reason || ''));
    failJobs = permanentFails.concat(retryResult.failJobs);

    okJobs = okJobs.concat(retryResult.okJobs);
    skipJobs = skipJobs.concat(retryResult.skipJobs);
    IMGCV_statsMerge_(stats, retryResult.stats);
  }

  const successJobs = okJobs.concat(skipJobs);
  const okCount = IMGCV_applyResults_(successJobs);

  // 今のリンク方針に合わせて自動判定
  try {
    const sRange = sheet.getRange(rowTop, IMGCV_COL_S, rows, 1);
    const tRange = sheet.getRange(rowTop, IMGCV_COL_T, rows, 1);

    if (typeof IMGCV_buildPlainRichText_ === 'function') {
      const sRich = S.map(row => [IMGCV_buildPlainRichText_(row[0])]);
      const tRich = T.map(row => [IMGCV_buildPlainRichText_(row[0])]);
      sRange.setRichTextValues(sRich);
      tRange.setRichTextValues(tRich);
    } else if (typeof IMGCV_buildRichWithLinks_ === 'function') {
      const sRich = S.map(row => [IMGCV_buildRichWithLinks_(row[0])]);
      const tRich = T.map(row => [IMGCV_buildRichWithLinks_(row[0])]);
      sRange.setRichTextValues(sRich);
      tRange.setRichTextValues(tRich);
    } else {
      sRange.setValues(S);
      tRange.setValues(T);
    }

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`[IMGCV] シート書き込みエラー: ${e}`);
    try {
      sheet.getRange(rowTop, IMGCV_COL_S, rows, 1).setValues(S);
      sheet.getRange(rowTop, IMGCV_COL_T, rows, 1).setValues(T);
      SpreadsheetApp.flush();
    } catch (e2) {
      Logger.log(`[IMGCV] フォールバック書き込みも失敗: ${e2}`);
    }
  }

  Logger.log(`[IMGCV][SLICE SUMMARY] jobs=${jobs.length} OK=${stats.ok} FAIL=${stats.fail} SKIP=${stats.skip} okApplied=${okCount}`);

  return {
    processedRows,
    okCount,
    stats,
    hadJobs: true
  };
}