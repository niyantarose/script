/******************************************************
 * Yahoo画像 全削除→全アップロード→後段submit
 * v33.0a: reupload_urls skip_submit 対応
 * - C列の商品コード名のDriveフォルダから画像取得
 * - S/T列URLに依存しない
 * - VPSローカルフォルダ不要
 * - 免責画像を末尾ソート、商品画像はサフィックス順
 * - it-07002 は submit リトライ除外
 * - シート反映は最後に1回、対象行のみ一括書込
 * - reupload / submit は波状並列で高速化
 * - 追加シートは作らない
 * - 実行中は console.log、最後に alert / toast で通知
 * - ボタン割り当て運用のため onOpen は使わない
 ******************************************************/

const 設定_ヤフー全削除再UP = {
  シート名: '①商品入力シート',
  ヘッダー行数: 2,
  列_商品コード: 3,
  列_メイン画像: 19,
  列_詳細画像: 20,
  DRIVE_IMAGE_FOLDER_ID: '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz',
  SELLER_ID: 'niyantarose',
  YAHOO_CDN_BASE: 'https://item-shopping.c.yimg.jp/i/n/',
  EXECUTE_API: 'https://img.niyantarose.com/yahoo_zip_upload.php',
  ZIP_BATCH_MAX_FILES: 6,
  REUPLOAD_WAVE_SIZE: 20,
  SUBMIT_CHUNK_SIZE: 200,
  SUBMIT_WAVE_SIZE: 20,
  SUBMIT_RETRY_WAITS_MS: [0, 5000, 10000, 15000, 20000, 25000],
  RESERVE_PUBLISH_RETRY_WAITS_MS: [0, 3000, 5000, 8000],
  DRIVE_FOLDER_QUERY_CHUNK: 40,
  DRIVE_FILE_PARENT_QUERY_CHUNK: 50,
  SHEET_WRITE_PLAIN_URL: false,
  NOTIFY_SAMPLE_MAX: 10,
  STATE_KEY: 'YAHOO_ALL_REUPLOAD_STATE_V330A',
  RUN_TOKEN_KEY: 'YAHOO_ALL_REUPLOAD_RUN_TOKEN_V2',
  VPS_SECRET_PROP_KEY: 'VPS_SECRET',
};

function API再アップロード_一括実行() { ヤフー全削除再UP_一括実行(); }
function API再アップロード_再開() { ヤフー全削除再UP_再開(); }
function API再アップロード_中断要求() { ヤフー全削除再UP_中断要求(); }
function API再アップロード_リセット() { ヤフー全削除再UP_リセット(); }

function ヤフー全削除再UP_一括実行() {
  ヤフー全削除再UP_リセット();
  const runToken = 実行トークン発行_();

  try {
    実行ログ_情報('一括実行 開始 v33.0a');

    const cfg = 設定_ヤフー全削除再UP;
    const sh = SpreadsheetApp.getActive().getSheetByName(cfg.シート名);
    if (!sh) throw new Error('シートが見つからない: ' + cfg.シート名);

    const 対象 = 対象ジョブ抽出_(sh);
    実行継続確認_({ runToken: runToken }, '対象抽出後');

    実行ログ_情報('対象抽出 完了', { 件数: 対象.length });
    if (対象.length === 0) {
      実行ログ_警告('対象0件');
      通知表示_({
        totalTargets: 0,
        successCount: 0,
        submitPendingCodes: [],
        deleteNgCodes: [],
        uploadNgCodes: [],
        submitNgCodes: [],
        finalSubmitNgCodes: [],
        unknownDeleteCodes: [],
        unknownUploadCodes: [],
        jsonErrorCodes: [],
        authMessages: [],
      }, null);
      実行トークン解放_({ runToken: runToken });
      return;
    }

    const バッチ群 = バッチ分割_(対象, cfg.ZIP_BATCH_MAX_FILES);
    const state = {
      phase: 'running',
      nextBatch: 0,
      batches: バッチ群,
      runToken: runToken,
      totalTargets: 対象.length,
      successCount: 0,
      submitPendingCodes: [],
      submitNgCodes: [],
      finalSubmitNgCodes: [],
      hardNgCodes: [],
      deleteNgCodes: [],
      uploadNgCodes: [],
      unknownDeleteCodes: [],
      unknownUploadCodes: [],
      jsonErrorCodes: [],
      authMessages: [],
      createdAt: new Date().toISOString(),
    };
    状態保存_(state);
    ヤフー全削除再UP_処理本体_(state);
  } catch (e) {
    if (中断エラーか_(e)) {
      実行ログ_警告('一括実行 中断', { reason: 中断理由_(e) });
      実行トークン解放_({ runToken: runToken });
      return;
    }
    throw e;
  }
}

function ヤフー全削除再UP_再開() {
  const st = 状態読込_();
  if (!st) {
    実行ログ_警告('再開する状態がありません');
    return;
  }

  if (!st.totalTargets) st.totalTargets = 0;
  if (!st.successCount) st.successCount = 0;
  if (!st.submitPendingCodes) st.submitPendingCodes = [];
  if (!st.submitNgCodes) st.submitNgCodes = [];
  if (!st.finalSubmitNgCodes) st.finalSubmitNgCodes = [];
  if (!st.hardNgCodes) st.hardNgCodes = [];
  if (!st.deleteNgCodes) st.deleteNgCodes = [];
  if (!st.uploadNgCodes) st.uploadNgCodes = [];
  if (!st.unknownDeleteCodes) st.unknownDeleteCodes = [];
  if (!st.unknownUploadCodes) st.unknownUploadCodes = [];
  if (!st.jsonErrorCodes) st.jsonErrorCodes = [];
  if (!st.authMessages) st.authMessages = [];

  st.runToken = 実行トークン発行_();

  try {
    実行ログ_情報('再開', {
      nextBatch: st.nextBatch,
      totalBatches: st.batches?.length,
      totalTargets: st.totalTargets,
      successCount: st.successCount,
      submitPendingCount: st.submitPendingCodes.length,
      submitNgCount: st.submitNgCodes.length,
      hardNgCount: st.hardNgCodes.length,
    });

    状態保存_(st);
    ヤフー全削除再UP_処理本体_(st);
  } catch (e) {
    if (中断エラーか_(e)) {
      実行ログ_警告('再開実行 中断', { reason: 中断理由_(e) });
      return;
    }
    throw e;
  }
}

function ヤフー全削除再UP_中断要求() {
  const stopToken = 'STOP_' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    設定_ヤフー全削除再UP.RUN_TOKEN_KEY,
    stopToken
  );
  実行ログ_警告('中断要求を受け付けました');

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '中断要求を受け付けました。現在の処理区切りで停止します',
      '中断要求',
      8
    );
  } catch (_) {}
}

function ヤフー全削除再UP_リセット() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(設定_ヤフー全削除再UP.STATE_KEY);
  props.deleteProperty(設定_ヤフー全削除再UP.RUN_TOKEN_KEY);
  実行ログ_情報('状態リセット');
}

function ヤフー全削除再UP_処理本体_(state) {
  const cfg = 設定_ヤフー全削除再UP;

  実行継続確認_(state, '処理開始');

  if (!state.totalTargets) state.totalTargets = 0;
  if (!state.successCount) state.successCount = 0;
  if (!state.submitPendingCodes) state.submitPendingCodes = [];
  if (!state.submitNgCodes) state.submitNgCodes = [];
  if (!state.finalSubmitNgCodes) state.finalSubmitNgCodes = [];
  if (!state.hardNgCodes) state.hardNgCodes = [];
  if (!state.deleteNgCodes) state.deleteNgCodes = [];
  if (!state.uploadNgCodes) state.uploadNgCodes = [];
  if (!state.unknownDeleteCodes) state.unknownDeleteCodes = [];
  if (!state.unknownUploadCodes) state.unknownUploadCodes = [];
  if (!state.jsonErrorCodes) state.jsonErrorCodes = [];
  if (!state.authMessages) state.authMessages = [];

  if (state.nextBatch < state.batches.length) {
    const remainingBatches = state.batches.slice(state.nextBatch);
    const secret = VPSシークレット取得_();
    const requests = remainingBatches.map(function(batch) {
      const items = batch.jobs.map(ジョブをitem変換_);
      return {
        url: cfg.EXECUTE_API,
        method: 'post',
        payload: {
          action: 'reupload_urls',
          seller_id: cfg.SELLER_ID,
          skip_submit: '1',
          items_json: JSON.stringify(items),
        },
        headers: { 'X-VPS-Secret': secret },
        muteHttpExceptions: true,
      };
    });

    const waveSize = Math.max(1, cfg.REUPLOAD_WAVE_SIZE || 20);
    実行ログ_情報('全バッチ並列reupload開始', { count: requests.length, waveSize: waveSize });
    const 成功ジョブ = [];

    for (let i = 0; i < requests.length; i += waveSize) {
      実行継続確認_(state, 'reupload wave ' + (Math.floor(i / waveSize) + 1) + ' 開始');
      const reqWave = requests.slice(i, i + waveSize);
      const respWave = UrlFetchApp.fetchAll(reqWave);
      実行継続確認_(state, 'reupload wave ' + (Math.floor(i / waveSize) + 1) + ' 応答後');

      respWave.forEach(function(resp, w) {
        const idx = i + w;
        const batch = remainingBatches[idx];
        const batchNo = state.nextBatch + idx + 1;
        const res = レスポンス要約_(resp, 'reupload_urls');

        const authMsgs = 認証エラーメッセージ抽出_(res);
        if (authMsgs.length > 0) {
          state.authMessages.push(...authMsgs);
          実行ログ_警告('バッチ' + batchNo + ' 認証エラー検出', { messages: authMsgs });
        }

        実行ログ_情報('バッチ' + batchNo + '結果', { ok: res.ok, http: res.http });

        if (!res.json) {
          const codes = batch.jobs.map(function(j) { return j.code; });
          実行ログ_警告('バッチ' + batchNo + ' 完全失敗（JSON応答なし）', { codes: codes });
          state.jsonErrorCodes.push(...codes);
          state.hardNgCodes.push(...codes);
          state.uploadNgCodes.push(...codes);
          return;
        }

        const judged = バッチ結果を行別判定_(res, batch, batchNo);

        if (judged.okJobs.length > 0) {
          成功ジョブ.push(...judged.okJobs);
          state.successCount += judged.okJobs.length;
        }

        if (judged.submitPendingCodes.length > 0) {
          state.submitPendingCodes.push(...judged.submitPendingCodes);
          実行ログ_情報('バッチ' + batchNo + ' submit保留', {
            count: judged.submitPendingCodes.length,
            codes: judged.submitPendingCodes,
          });
        }

        if (judged.submitNgCodes.length > 0) {
          state.submitNgCodes.push(...judged.submitNgCodes);
          実行ログ_警告('バッチ' + batchNo + ' submitNG', {
            count: judged.submitNgCodes.length,
            codes: judged.submitNgCodes,
          });
        }

        if (judged.deleteNgCodes.length > 0) {
          state.deleteNgCodes.push(...judged.deleteNgCodes);
          実行ログ_警告('バッチ' + batchNo + ' deleteNG', {
            count: judged.deleteNgCodes.length,
            codes: judged.deleteNgCodes,
          });
        }

        if (judged.uploadNgCodes.length > 0) {
          state.uploadNgCodes.push(...judged.uploadNgCodes);
          実行ログ_警告('バッチ' + batchNo + ' uploadNG', {
            count: judged.uploadNgCodes.length,
            codes: judged.uploadNgCodes,
          });
        }

        if (judged.hardNgCodes.length > 0) {
          state.hardNgCodes.push(...judged.hardNgCodes);
          実行ログ_警告('バッチ' + batchNo + ' delete/upload NG', {
            count: judged.hardNgCodes.length,
            codes: judged.hardNgCodes,
          });
        }

        if (judged.unknownDeleteCodes.length > 0) {
          state.unknownDeleteCodes.push(...judged.unknownDeleteCodes);
          実行ログ_警告('バッチ' + batchNo + ' delete詳細なし', {
            count: judged.unknownDeleteCodes.length,
            codes: judged.unknownDeleteCodes,
          });
        }

        if (judged.unknownUploadCodes.length > 0) {
          state.unknownUploadCodes.push(...judged.unknownUploadCodes);
          実行ログ_警告('バッチ' + batchNo + ' upload詳細なし', {
            count: judged.unknownUploadCodes.length,
            codes: judged.unknownUploadCodes,
          });
        }

        if (!res.ok) {
          実行ログ_警告('バッチ' + batchNo + ' API失敗', { http: res.http, bodyTail: res.body_tail });
        }
      });
    }

    実行継続確認_(state, 'CDN URL書き戻し前');
    if (成功ジョブ.length > 0) {
      実行ログ_情報('CDN URL書き戻し開始', { rows: 成功ジョブ.length });
      シート反映_CDN_(成功ジョブ);
      実行ログ_情報('CDN URL書き戻し完了');
    }

    state.submitPendingCodes = 配列ユニーク_(state.submitPendingCodes);
    state.submitNgCodes = 配列ユニーク_(state.submitNgCodes);
    state.nextBatch = state.batches.length;
    状態保存_(state);
  }

  実行継続確認_(state, 'submit初回実行前');

  const uniquePendingCodes = 配列ユニーク_(state.submitPendingCodes);
  let finalSubmitNgCodes = [];

  if (uniquePendingCodes.length > 0) {
    実行ログ_情報('submit初回実行開始', { count: uniquePendingCodes.length });

    const firstSubmitRes = VPS_submitOnly_並列_(uniquePendingCodes);
    const firstAuthMsgs = 認証エラーメッセージ抽出_(firstSubmitRes);
    if (firstAuthMsgs.length > 0) {
      state.authMessages.push(...firstAuthMsgs);
      実行ログ_警告('submit初回実行 認証エラー検出', { messages: firstAuthMsgs });
    }

    const firstOkCount = firstSubmitRes.json?.submit?.ok_count || 0;
    const firstNgCount = firstSubmitRes.json?.submit?.ng_count || 0;
    実行ログ_情報('submit初回実行結果', {
      ok: firstSubmitRes.ok,
      ok_count: firstOkCount,
      ng_count: firstNgCount,
    });

    const firstNg = submitNGコード抽出_(firstSubmitRes, uniquePendingCodes);
    state.submitNgCodes = firstNg.allNgCodes;
    状態保存_(state);

    let retryFails = firstNg.retryableNgCodes;
    if (retryFails.length > 0) {
      実行ログ_情報('submitリトライ開始', { count: retryFails.length });
    }

    for (let r = 1; r < cfg.SUBMIT_RETRY_WAITS_MS.length; r++) {
      実行継続確認_(state, 'submit retry ' + r + ' 開始');
      if (retryFails.length === 0) break;

      const waitMs = cfg.SUBMIT_RETRY_WAITS_MS[r] || 0;
      実行ログ_情報('submitリトライ ' + r + '回目', { count: retryFails.length, wait_ms: waitMs });
      if (waitMs > 0) Utilities.sleep(waitMs);

      const retryRes = VPS_submitOnly_並列_(retryFails);
      const authMsgs = 認証エラーメッセージ抽出_(retryRes);
      if (authMsgs.length > 0) {
        state.authMessages.push(...authMsgs);
        実行ログ_警告('submitリトライ' + r + ' 認証エラー検出', { messages: authMsgs });
        break;
      }

      const ok2 = retryRes.json?.submit?.ok_count || 0;
      const ng2 = retryRes.json?.submit?.ng_count || 0;
      実行ログ_情報('submitリトライ' + r + '結果', { ok: retryRes.ok, ok_count: ok2, ng_count: ng2 });

      const retryNg = submitNGコード抽出_(retryRes, retryFails);
      retryFails = retryNg.retryableNgCodes;
    }

    finalSubmitNgCodes = retryFails;
  }

  state.finalSubmitNgCodes = 配列ユニーク_(finalSubmitNgCodes);
  状態保存_(state);

  実行継続確認_(state, 'reservePublish前');

  実行ログ_情報('ストアの全反映（reservePublish）を開始します');
  const waits = cfg.RESERVE_PUBLISH_RETRY_WAITS_MS || [0, 3000, 5000, 8000];

  let pubRes = null;
  for (let i = 0; i < waits.length; i++) {
    実行継続確認_(state, 'reservePublish retry ' + (i + 1) + ' 開始');

    const w = waits[i] || 0;
    if (w > 0) Utilities.sleep(w);

    pubRes = VPS_reservePublish_();

    const authMsgs = 認証エラーメッセージ抽出_(pubRes);
    if (authMsgs.length > 0) {
      state.authMessages.push(...authMsgs);
      実行ログ_警告('reservePublish 認証エラー検出', { messages: authMsgs });
      break;
    }

    if (pubRes.ok) break;

    const busy = String(pubRes.body_tail || '').includes('ed-00006');
    if (!busy) break;

    実行ログ_情報('Yahooがまだ処理中。reservePublish再試行待機', {
      try: i + 1,
      next_wait_ms: waits[i + 1] || 0,
    });
  }

  if (pubRes && pubRes.ok) {
    実行ログ_情報('全反映予約に成功しました', { 予約時間: pubRes.json?.reserve_time });
  } else {
    実行ログ_警告('全反映予約に失敗しました', { http: pubRes?.http, エラー内容: pubRes?.body_tail });
  }

  const hardNgUnique = 配列ユニーク_(state.hardNgCodes || []);
  if (hardNgUnique.length > 0) {
    実行ログ_警告('delete/upload NGあり', {
      count: hardNgUnique.length,
      sample: hardNgUnique.slice(0, 30),
    });
  }

  通知表示_(state, pubRes);

  PropertiesService.getScriptProperties().deleteProperty(設定_ヤフー全削除再UP.STATE_KEY);
  実行トークン解放_(state);
  実行ログ_情報('全処理完了');

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Yahoo画像処理が完了しました', '完了', 8);
  } catch (_) {}
}

function 実行トークン発行_() {
  const token = Utilities.getUuid() + '_' + Date.now();
  PropertiesService.getScriptProperties().setProperty(
    設定_ヤフー全削除再UP.RUN_TOKEN_KEY,
    token
  );
  return token;
}

function 実行トークン現在値_() {
  return PropertiesService.getScriptProperties().getProperty(
    設定_ヤフー全削除再UP.RUN_TOKEN_KEY
  ) || '';
}

function 実行継続確認_(state, point) {
  const expected = String(state?.runToken || '');
  const current = 実行トークン現在値_();

  if (!expected || current !== expected) {
    const where = point ? ' (' + point + ')' : '';
    throw new Error('RUN_ABORTED: より新しい実行または中断要求を検出したため停止しました' + where);
  }
}

function 実行トークン解放_(state) {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty(設定_ヤフー全削除再UP.RUN_TOKEN_KEY) || '';
  const expected = String(state?.runToken || '');

  if (expected && current === expected) {
    props.deleteProperty(設定_ヤフー全削除再UP.RUN_TOKEN_KEY);
  }
}

function 中断エラーか_(e) {
  const msg = String(e?.message || e || '');
  return msg.indexOf('RUN_ABORTED:') === 0;
}

function 中断理由_(e) {
  const msg = String(e?.message || e || '');
  return msg.replace(/^RUN_ABORTED:\s*/, '');
}

function 対象ジョブ抽出_(sh) {
  const cfg = 設定_ヤフー全削除再UP;
  const startRow = cfg.ヘッダー行数 + 1;
  const lastRow = sh.getLastRow();
  const num = Math.max(0, lastRow - startRow + 1);
  if (num === 0) return [];

  const rows = sh.getRange(startRow, cfg.列_商品コード, num, 1).getDisplayValues();
  const codes = [];
  const rowIndexMap = {};

  for (let i = 0; i < rows.length; i++) {
    const code = String(rows[i][0] || '').trim();
    if (!code) continue;
    if (!rowIndexMap[code]) {
      rowIndexMap[code] = startRow + i;
      codes.push(code);
    }
  }

  if (codes.length === 0) return [];

  実行ログ_情報('Drive対象フォルダ検索中...', { codeCount: codes.length });
  const folderMap = Drive_対象フォルダマップ取得_(
    cfg.DRIVE_IMAGE_FOLDER_ID,
    codes,
    cfg.DRIVE_FOLDER_QUERY_CHUNK
  );
  実行ログ_情報('Drive対象フォルダ検索完了', { hitCount: Object.keys(folderMap).length });

  const targetFolderIds = [];
  const codeToFolderId = {};
  for (const code of codes) {
    const folderId = folderMap[code];
    if (!folderId) {
      実行ログ_警告('Driveフォルダなし', { code: code });
      continue;
    }
    codeToFolderId[code] = folderId;
    targetFolderIds.push(folderId);
  }

  if (targetFolderIds.length === 0) return [];

  実行ログ_情報('Drive画像ファイル一括取得中...', { folderCount: targetFolderIds.length });
  const filesByFolder = Drive_フォルダ群ファイル一括取得_(targetFolderIds);

  const rawJobs = [];
  for (const code of codes) {
    const folderId = codeToFolderId[code];
    if (!folderId) continue;

    const files = filesByFolder[folderId] || [];
    if (files.length === 0) {
      実行ログ_警告('画像ファイルなし', { code: code });
      continue;
    }

    const sources = files.map(function(f) {
      return { type: 'drive', fileId: f.id, fileName: f.name };
    });
    rawJobs.push({
      code: code,
      sources: ソース免責ソート_(sources),
      rowIndex: rowIndexMap[code],
    });
  }

  実行ログ_情報('対象抽出 完了', { 件数: rawJobs.length });
  return rawJobs;
}

function ジョブをitem変換_(job) {
  function srcToUrl(src) {
    if (src.type === 'drive') {
      return 'https://drive.google.com/uc?export=download&id=' + src.fileId;
    }
    return src.url || '';
  }

  const s = job.sources.length > 0 ? srcToUrl(job.sources[0]) : '';
  const tUrls = job.sources.slice(1).map(srcToUrl).filter(Boolean);

  return {
    item_code: job.code,
    s: s,
    t: tUrls.join(';'),
  };
}

function バッチ分割_(jobs, maxFiles) {
  const out = [];
  const n = Math.max(1, Number(maxFiles) || 1);
  for (let i = 0; i < jobs.length; i += n) {
    out.push({ jobs: jobs.slice(i, i + n) });
  }
  return out;
}

function Drive_対象フォルダマップ取得_(parentFolderId, codes, chunkSize) {
  if (!Drive || !Drive.Files || !Drive.Files.list) {
    throw new Error('Drive API（高度なGoogleサービス）を有効化してください。');
  }

  const uniqCodes = [...new Set((codes || []).map(function(v) {
    return String(v || '').trim();
  }).filter(Boolean))];
  const out = {};
  if (uniqCodes.length === 0) return out;

  const CHUNK = Math.max(1, chunkSize || 40);
  for (let i = 0; i < uniqCodes.length; i += CHUNK) {
    const chunk = uniqCodes.slice(i, i + CHUNK);
    const nameOr = chunk.map(function(code) {
      return "name='" + Drive_クエリエスケープ_(code) + "'";
    }).join(' or ');
    const q = "'" + parentFolderId + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false and (" + nameOr + ")";

    let pageToken = null;
    do {
      const res = Drive.Files.list({
        q: q,
        fields: 'nextPageToken, files(id,name)',
        pageSize: 1000,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });

      (res.files || []).forEach(function(f) {
        if (!out[f.name]) out[f.name] = f.id;
      });

      pageToken = res.nextPageToken;
    } while (pageToken);
  }

  return out;
}

function Drive_フォルダ群ファイル一括取得_(folderIds) {
  if (!Drive || !Drive.Files || !Drive.Files.list) {
    throw new Error('Drive API（高度なGoogleサービス）を有効化してください。');
  }

  const byFolder = {};
  (folderIds || []).forEach(function(id) { byFolder[id] = []; });

  const CHUNK = Math.max(1, 設定_ヤフー全削除再UP.DRIVE_FILE_PARENT_QUERY_CHUNK || 50);
  for (let i = 0; i < folderIds.length; i += CHUNK) {
    const chunk = folderIds.slice(i, i + CHUNK);
    const parentOr = chunk.map(function(id) {
      return "'" + id + "' in parents";
    }).join(' or ');
    const q = '(' + parentOr + ") and mimeType contains 'image/' and trashed=false";

    let pageToken = null;
    do {
      const res = Drive.Files.list({
        q: q,
        fields: 'nextPageToken, files(id,name,parents)',
        pageSize: 1000,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });

      for (const f of (res.files || [])) {
        for (const p of (f.parents || [])) {
          if (byFolder[p]) byFolder[p].push({ id: f.id, name: f.name });
        }
      }

      pageToken = res.nextPageToken;
    } while (pageToken);
  }

  return byFolder;
}

function Drive_クエリエスケープ_(s) {
  return String(s || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function ソース免責ソート_(sources) {
  const product = [];
  const disclaimer = [];

  for (const src of sources) {
    const name = String(src.fileName || '').toLowerCase();
    const isDisc = name.startsWith('notice_') || name.includes('免責');
    if (isDisc) disclaimer.push(src);
    else product.push(src);
  }

  product.sort(function(a, b) {
    return _サフィックス番号取得(a.fileName) - _サフィックス番号取得(b.fileName);
  });

  return [...product, ...disclaimer];
}

function _サフィックス番号取得(fileName) {
  if (!fileName) return 999;
  const base = String(fileName).replace(/\.[^.]+$/, '');
  const m = base.match(/_(\d+)$/);
  if (m) return parseInt(m[1], 10);
  return 0;
}

function シート反映_CDN_(jobs) {
  const cfg = 設定_ヤフー全削除再UP;
  const sh = SpreadsheetApp.getActive().getSheetByName(cfg.シート名);
  if (!sh || !jobs || jobs.length === 0) return;

  const rowMap = {};
  for (const job of jobs) {
    if (job && job.rowIndex) rowMap[job.rowIndex] = job;
  }

  const rows = Object.keys(rowMap).map(Number).sort(function(a, b) { return a - b; });
  if (rows.length === 0) return;

  let seg = [rows[0]];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === rows[i - 1] + 1) {
      seg.push(rows[i]);
    } else {
      _書き込みセグメント(sh, seg, rowMap, cfg);
      seg = [rows[i]];
    }
  }
  _書き込みセグメント(sh, seg, rowMap, cfg);
}

function _書き込みセグメント(sh, rows, rowMap, cfg) {
  if (!rows || rows.length === 0) return;

  if (cfg.SHEET_WRITE_PLAIN_URL) {
    const values = rows.map(function(rowNum) {
      const job = rowMap[rowNum];
      const mainUrl = cfg.YAHOO_CDN_BASE + cfg.SELLER_ID + '_' + job.code;
      const details = [];
      for (let i = 1; i < job.sources.length; i++) {
        details.push(cfg.YAHOO_CDN_BASE + cfg.SELLER_ID + '_' + job.code + '_' + i);
      }
      return [mainUrl, details.join('\n')];
    });
    sh.getRange(rows[0], cfg.列_メイン画像, rows.length, 2).setValues(values);
    return;
  }

  const values = rows.map(function(rowNum) {
    const job = rowMap[rowNum];
    const mainUrl = cfg.YAHOO_CDN_BASE + cfg.SELLER_ID + '_' + job.code;
    const mainRich = SpreadsheetApp.newRichTextValue()
      .setText(mainUrl)
      .setLinkUrl(mainUrl)
      .build();

    const details = [];
    for (let i = 1; i < job.sources.length; i++) {
      details.push(cfg.YAHOO_CDN_BASE + cfg.SELLER_ID + '_' + job.code + '_' + i);
    }

    const detailRich = 詳細URLリッチ作成_(details);
    return [mainRich, detailRich];
  });

  sh.getRange(rows[0], cfg.列_メイン画像, rows.length, 2).setRichTextValues(values);
}

function 詳細URLリッチ作成_(urls) {
  if (!urls || urls.length === 0) {
    return SpreadsheetApp.newRichTextValue().setText('').build();
  }

  const fullText = urls.join('\n');
  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  let pos = 0;
  for (const url of urls) {
    builder.setLinkUrl(pos, pos + url.length, url);
    pos += url.length + 1;
  }
  return builder.build();
}

function VPS_submitOnly_(itemCodes) {
  const cfg = 設定_ヤフー全削除再UP;
  const secret = VPSシークレット取得_();

  const resp = UrlFetchApp.fetch(cfg.EXECUTE_API, {
    method: 'post',
    payload: {
      action: 'submit',
      seller_id: cfg.SELLER_ID,
      item_codes: JSON.stringify(itemCodes),
    },
    headers: { 'X-VPS-Secret': secret },
    muteHttpExceptions: true,
  });

  return レスポンス要約_(resp, 'submit');
}

function VPS_submitOnly_並列_(itemCodes) {
  const cfg = 設定_ヤフー全削除再UP;
  const secret = VPSシークレット取得_();
  const codes = 配列ユニーク_(itemCodes);

  if (codes.length === 0) {
    return {
      ok: true,
      http: 200,
      body_tail: '',
      json: {
        ok: true,
        submit: { ok_count: 0, ng_count: 0, details: [] },
      },
    };
  }

  const chunkSize = Math.max(1, cfg.SUBMIT_CHUNK_SIZE || 200);
  const waveSize = Math.max(1, cfg.SUBMIT_WAVE_SIZE || 20);
  const reqs = [];

  for (let i = 0; i < codes.length; i += chunkSize) {
    const chunk = codes.slice(i, i + chunkSize);
    reqs.push({
      url: cfg.EXECUTE_API,
      method: 'post',
      payload: {
        action: 'submit',
        seller_id: cfg.SELLER_ID,
        item_codes: JSON.stringify(chunk),
      },
      headers: { 'X-VPS-Secret': secret },
      muteHttpExceptions: true,
    });
  }

  let allOk = true;
  let maxHttp = 200;
  let bodyTail = '';
  let totalOk = 0;
  let totalNg = 0;
  let firstMessage = '';
  let firstError = '';
  let firstTokenDebug = null;
  const details = [];

  for (let i = 0; i < reqs.length; i += waveSize) {
    const respWave = UrlFetchApp.fetchAll(reqs.slice(i, i + waveSize));
    for (const resp of respWave) {
      const r = レスポンス要約_(resp, 'submit');
      if (!r.ok) allOk = false;
      if (r.http > maxHttp) maxHttp = r.http;
      if (r.body_tail) bodyTail = r.body_tail;
      if (!firstMessage && r.json?.message) firstMessage = r.json.message;
      if (!firstError && r.json?.error) firstError = r.json.error;
      if (!firstTokenDebug && r.json?.token_debug) firstTokenDebug = r.json.token_debug;

      const s = r.json?.submit;
      if (!s) continue;
      totalOk += s.ok_count || 0;
      totalNg += s.ng_count || 0;
      if (Array.isArray(s.details)) details.push(...s.details);
    }
  }

  return {
    ok: allOk,
    http: maxHttp,
    body_tail: bodyTail,
    json: {
      ok: allOk,
      message: firstMessage,
      error: firstError,
      token_debug: firstTokenDebug,
      submit: {
        ok_count: totalOk,
        ng_count: totalNg,
        details: details,
      },
    },
  };
}

function VPS_reservePublish_() {
  const cfg = 設定_ヤフー全削除再UP;
  const secret = VPSシークレット取得_();

  const resp = UrlFetchApp.fetch(cfg.EXECUTE_API, {
    method: 'post',
    payload: {
      action: 'reserve_publish',
      seller_id: cfg.SELLER_ID,
    },
    headers: { 'X-VPS-Secret': secret },
    muteHttpExceptions: true,
  });

  return レスポンス要約_(resp, 'reserve_publish');
}

function 状態保存_(state) {
  const props = PropertiesService.getScriptProperties();

  if (state && state.runToken) {
    const current = props.getProperty(設定_ヤフー全削除再UP.RUN_TOKEN_KEY) || '';
    if (current !== state.runToken) {
      実行ログ_警告('状態保存スキップ: runToken不一致', {
        expected: state.runToken,
        current: current,
      });
      return false;
    }
  }

  props.setProperty(
    設定_ヤフー全削除再UP.STATE_KEY,
    JSON.stringify(state)
  );
  return true;
}

function 状態読込_() {
  const s = PropertiesService.getScriptProperties().getProperty(設定_ヤフー全削除再UP.STATE_KEY);
  return s ? JSON.parse(s) : null;
}

function VPSシークレット取得_() {
  const cfg = 設定_ヤフー全削除再UP;
  const v = PropertiesService.getScriptProperties().getProperty(cfg.VPS_SECRET_PROP_KEY);
  if (!v) throw new Error('VPSシークレット未設定。ScriptPropertiesに ' + cfg.VPS_SECRET_PROP_KEY + ' をセットしてください。');
  return String(v).trim();
}

function レスポンス要約_(resp, tag) {
  const http = resp.getResponseCode();
  const text = resp.getContentText() || '';
  const body_tail = text.slice(Math.max(0, text.length - 400));

  let json = null;
  const t = text.trim();
  if (t && (t[0] === '{' || t[0] === '[')) {
    try {
      json = JSON.parse(t);
    } catch (_) {}
  }

  const ok = (http >= 200 && http < 300) && (!json || json.ok !== false);
  if (!ok) {
    実行ログ_警告('API失敗', { tag: tag, http: http, bodyTail: body_tail, jsonOk: json?.ok });
  }

  return { ok: ok, http: http, body_tail: body_tail, json: json };
}

function 実行ログ_情報(msg, obj) { 実行ログ_('INFO', msg, obj); }
function 実行ログ_警告(msg, obj) { 実行ログ_('WARN', msg, obj); }
function 実行ログ_エラー(msg, obj) { 実行ログ_('ERROR', msg, obj); }

function 実行ログ_(level, msg, obj) {
  const ts = new Date().toISOString();
  const line = obj
    ? '[' + ts + '] [' + level + '] ' + msg + ' ' + JSON.stringify(obj)
    : '[' + ts + '] [' + level + '] ' + msg;
  console.log(line);
}

function バッチ結果を行別判定_(res, batch, batchNo) {
  const json = res.json || {};
  const deleteKnown = セクション有無_(json, 'delete');
  const uploadKnown = セクション有無_(json, 'upload');
  const submitKnown = セクション有無_(json, 'submit');
  const submitSkipped = !!json?.submit?.skipped;
  const pendingSubmitSet = 配列をSet_(json?.submit?.pending_item_codes);

  const deleteMap = 詳細配列をコードMap_(結果詳細配列取得_(json, 'delete'));
  const uploadMap = 詳細配列をコードMap_(結果詳細配列取得_(json, 'upload'));
  const submitMap = 詳細配列をコードMap_(結果詳細配列取得_(json, 'submit'));

  const okJobs = [];
  const submitPendingCodes = [];
  const submitNgCodes = [];
  const deleteNgCodes = [];
  const uploadNgCodes = [];
  const hardNgCodes = [];
  const unknownDeleteCodes = [];
  const unknownUploadCodes = [];

  batch.jobs.forEach(function(job) {
    const d = deleteMap[job.code];
    const u = uploadMap[job.code];
    const s = submitMap[job.code];

    const deleteOk = deleteKnown ? 詳細OK判定_(d, false) : null;
    const uploadOk = uploadKnown ? 詳細OK判定_(u, false) : null;
    const submitOk = submitKnown && !submitSkipped ? 詳細OK判定_(s, false) : null;

    const deletePass = deleteKnown ? deleteOk === true : res.ok;
    const uploadPass = uploadKnown ? uploadOk === true : res.ok;
    const phase1Ok = deletePass && uploadPass;
    const pendingSubmit = submitSkipped
      ? (pendingSubmitSet.size > 0 ? pendingSubmitSet.has(job.code) : phase1Ok)
      : false;
    const overallOk = submitSkipped
      ? phase1Ok && pendingSubmit
      : (phase1Ok && (submitKnown ? submitOk === true : res.ok));

    if (overallOk) okJobs.push(job);
    if (pendingSubmit) submitPendingCodes.push(job.code);

    if (!deleteKnown) unknownDeleteCodes.push(job.code);
    if (!uploadKnown) unknownUploadCodes.push(job.code);
    if (deleteKnown && deleteOk === false) deleteNgCodes.push(job.code);
    if (uploadKnown && uploadOk === false) uploadNgCodes.push(job.code);
    if (!submitSkipped && submitKnown && submitOk === false) submitNgCodes.push(job.code);
    if ((deleteKnown && deleteOk === false) || (uploadKnown && uploadOk === false)) {
      hardNgCodes.push(job.code);
    }
  });

  実行ログ_情報('バッチ' + batchNo + '判定', {
    okJobs: okJobs.length,
    submitPending: submitPendingCodes.length,
    deleteNg: deleteNgCodes.length,
    uploadNg: uploadNgCodes.length,
    submitNg: submitNgCodes.length,
    unknownDelete: unknownDeleteCodes.length,
    unknownUpload: unknownUploadCodes.length,
  });

  return {
    okJobs: okJobs,
    submitPendingCodes: submitPendingCodes,
    submitNgCodes: submitNgCodes,
    deleteNgCodes: deleteNgCodes,
    uploadNgCodes: uploadNgCodes,
    hardNgCodes: hardNgCodes,
    unknownDeleteCodes: unknownDeleteCodes,
    unknownUploadCodes: unknownUploadCodes,
  };
}

function 結果詳細配列取得_(json, key) {
  const candidates = [
    json?.[key]?.details,
    json?.[key + '_details'],
    json?.details?.[key],
    json?.results?.[key],
    json?.[key]?.items,
  ];
  for (const v of candidates) {
    if (Array.isArray(v)) return v;
  }
  return [];
}

function セクション有無_(json, key) {
  return Array.isArray(json?.[key]?.details)
      || Array.isArray(json?.[key + '_details'])
      || Array.isArray(json?.details?.[key])
      || Array.isArray(json?.results?.[key])
      || Array.isArray(json?.[key]?.items)
      || !!json?.[key];
}

function 詳細配列をコードMap_(details) {
  const map = {};
  (details || []).forEach(function(d) {
    const code = String(d?.item_code || d?.code || d?.itemCode || '').trim();
    if (code) map[code] = d;
  });
  return map;
}

function 詳細OK判定_(detail, fallback) {
  if (!detail) return fallback;
  if (typeof detail.ok === 'boolean') return detail.ok;
  if (typeof detail.success === 'boolean') return detail.success;
  if (typeof detail.result === 'string') return /^(ok|success)$/i.test(detail.result);
  return fallback;
}

function 配列ユニーク_(arr) {
  return [...new Set((arr || []).map(function(v) {
    return String(v || '').trim();
  }).filter(Boolean))];
}

function 配列をSet_(arr) {
  return new Set(配列ユニーク_(Array.isArray(arr) ? arr : []));
}

function サンプル文字列_(arr, limit) {
  const uniq = 配列ユニーク_(arr);
  if (uniq.length === 0) return '';
  return uniq.slice(0, limit).join(', ');
}

function submitNGコード抽出_(res, requestedCodes) {
  const requested = 配列ユニーク_(requestedCodes);
  const details = res?.json?.submit?.details;

  if (!Array.isArray(details) || details.length === 0) {
    const ng = res?.ok ? [] : requested;
    return {
      allNgCodes: ng,
      retryableNgCodes: ng,
    };
  }

  const byCode = 詳細配列をコードMap_(details);
  const allNgCodes = [];
  const retryableNgCodes = [];

  requested.forEach(function(code) {
    const detail = byCode[code];
    const ok = detail ? 詳細OK判定_(detail, false) === true : false;
    if (ok) return;

    allNgCodes.push(code);

    const bodyTail = String(detail?.body_tail || '');
    if (bodyTail.indexOf('it-07002') === -1) {
      retryableNgCodes.push(code);
    }
  });

  return {
    allNgCodes: 配列ユニーク_(allNgCodes),
    retryableNgCodes: 配列ユニーク_(retryableNgCodes),
  };
}

function 認証エラーメッセージ抽出_(res) {
  const out = [];
  const json = res?.json || {};
  const texts = [];

  function pushText(v) {
    if (v === null || v === undefined) return;
    const s = String(v).trim();
    if (s) texts.push(s);
  }

  pushText(json?.message);
  pushText(json?.error);
  pushText(json?.token_debug?.token_error_body_tail);
  pushText(res?.body_tail);

  ['delete', 'upload', 'submit'].forEach(function(key) {
    const details = json?.[key]?.details || [];
    details.forEach(function(d) {
      pushText(d?.message);
      pushText(d?.body_tail);
      pushText(d?.error_code);
    });
  });

  const merged = texts.join('\n');
  if (/consent has been revoked/i.test(merged)) {
    out.push('Yahoo認証が失効しています。VPSの refresh token を再発行してください');
  } else if (/invalid_grant/i.test(merged)) {
    out.push('Yahoo認証に失敗しました。VPSの refresh token を確認してください');
  } else if (/access_token missing and refresh failed/i.test(merged)) {
    out.push('Yahoo認証に失敗しました。VPSの access token / refresh token を確認してください');
  } else if (/invalid_token|expired_token|token expired|access token expired/i.test(merged)) {
    out.push('Yahooの access token の更新に失敗しました。VPS側の認証状態を確認してください');
  } else if ((res?.http || 0) === 401) {
    out.push('Yahoo認証エラーです。VPSの access token / refresh token を確認してください');
  }

  return [...new Set(out)];
}

function 通知メッセージ作成_(state, pubRes) {
  const cfg = 設定_ヤフー全削除再UP;
  const submitPending = 配列ユニーク_(state.submitPendingCodes);
  const deleteNg = 配列ユニーク_(state.deleteNgCodes);
  const uploadNg = 配列ユニーク_(state.uploadNgCodes);
  const submitNg = 配列ユニーク_(state.submitNgCodes);
  const finalSubmitNg = 配列ユニーク_(state.finalSubmitNgCodes);
  const unknownDelete = 配列ユニーク_(state.unknownDeleteCodes);
  const unknownUpload = 配列ユニーク_(state.unknownUploadCodes);
  const jsonErrors = 配列ユニーク_(state.jsonErrorCodes);
  const authMessages = 配列ユニーク_(state.authMessages);

  const lines = [];
  lines.push('Yahoo画像処理 完了');
  lines.push('対象: ' + (state.totalTargets || 0) + '件');
  lines.push('CDN反映: ' + (state.successCount || 0) + '件');
  lines.push('submit対象: ' + submitPending.length + '件');

  if (authMessages.length > 0) {
    lines.push('認証エラー: ' + authMessages.length + '件');
    authMessages.slice(0, 3).forEach(function(m) { lines.push('  ' + m); });
  }

  if (deleteNg.length > 0) {
    lines.push('delete NG: ' + deleteNg.length + '件');
    lines.push('  例: ' + サンプル文字列_(deleteNg, cfg.NOTIFY_SAMPLE_MAX));
  }

  if (uploadNg.length > 0) {
    lines.push('upload NG: ' + uploadNg.length + '件');
    lines.push('  例: ' + サンプル文字列_(uploadNg, cfg.NOTIFY_SAMPLE_MAX));
  }

  if (submitNg.length > 0) {
    lines.push('submit 初回NG: ' + submitNg.length + '件');
    lines.push('  例: ' + サンプル文字列_(submitNg, cfg.NOTIFY_SAMPLE_MAX));
  }

  if (finalSubmitNg.length > 0) {
    lines.push('submit 最終NG: ' + finalSubmitNg.length + '件');
    lines.push('  例: ' + サンプル文字列_(finalSubmitNg, cfg.NOTIFY_SAMPLE_MAX));
  } else if (submitNg.length > 0) {
    lines.push('submit 最終NG: 0件（リトライで解消）');
  }

  if (unknownDelete.length > 0) {
    lines.push('delete詳細不明: ' + unknownDelete.length + '件');
    lines.push('  例: ' + サンプル文字列_(unknownDelete, cfg.NOTIFY_SAMPLE_MAX));
  }

  if (unknownUpload.length > 0) {
    lines.push('upload詳細不明: ' + unknownUpload.length + '件');
    lines.push('  例: ' + サンプル文字列_(unknownUpload, cfg.NOTIFY_SAMPLE_MAX));
  }

  if (jsonErrors.length > 0) {
    lines.push('JSON応答なし: ' + jsonErrors.length + '件');
    lines.push('  例: ' + サンプル文字列_(jsonErrors, cfg.NOTIFY_SAMPLE_MAX));
  }

  lines.push('reservePublish: ' + (pubRes?.ok ? '成功' : '失敗'));
  return lines.join('\n');
}

function 通知表示_(state, pubRes) {
  const msg = 通知メッセージ作成_(state, pubRes);
  実行ログ_情報('最終通知', { message: msg });

  try {
    SpreadsheetApp.getUi().alert(msg);
    return;
  } catch (_) {}

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Yahoo画像処理結果', 15);
  } catch (_) {}

  console.log(msg);
}
