// background.js - ダウンロード制御サービスワーカー

const GAS_SYNC_QUEUE_KEY = 'booksGasSyncQueueV1';
const GAS_SYNC_STATUS_KEY = 'booksGasSyncStatusV1';
const GAS_SYNC_RESET_TOKEN_KEY = 'booksGasSyncResetTokenV1';
const GAS_SYNC_ALARM_NAME = 'booksGasSyncQueueAlarmV1';
const GAS_SYNC_TIMEOUT_MS = 90000;
// 一時的な通信失敗・タイムアウト・コールドスタートを吸収するための自動リトライ設定。
// 1回目で失敗しても最大この回数まで同一ジョブを再送する（指数バックオフ）。
const GAS_SYNC_MAX_ATTEMPTS = 3;
const GAS_SYNC_RETRY_BASE_DELAY_MS = 3000;
const GAS_SYNC_ALARM_PERIOD_MINUTES = 1;
const IMAGE_FILENAME_EXT = 'jpg';
const TAIWAN_IMAGE_MAX_EDGE = 1000;
const TAIWAN_DOWNLOAD_MARKER = 'nyantarose_tw___-_';
const pendingTaiwanFilenameByMarker = new Map();
const TAIWAN_FILENAME_HOOK_TTL_MS = 30000;
let taiwanFilenameHookCleanupTimer = null;
let taiwanFilenameHookRegistered = false;
let gasSyncProcessing = false;
let gasSyncResetToken = 0;


function storageGetLocal(keys) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(keys, result => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message || 'chrome.storage.local.get failed'));
        return;
      }
      resolve(result || {});
    });
  });
}

function storageSetLocal(items) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(items, () => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message || 'chrome.storage.local.set failed'));
        return;
      }
      resolve();
    });
  });
}

function buildGasSyncStatus(status) {
  return {
    state: '',
    jobId: '',
    itemCount: 0,
    queueLength: 0,
    message: '',
    error: '',
    appended: 0,
    requestedAt: '',
    startedAt: '',
    finishedAt: '',
    ...status,
  };
}

function parseIsoTimeMs(value) {
  const time = Date.parse(String(value || ''));
  return Number.isFinite(time) ? time : 0;
}

function isGasSyncStatusStale(status) {
  if (!status?.state) return false;
  if (status.state === 'running') {
    const startedAtMs = parseIsoTimeMs(status.startedAt || status.requestedAt);
    return startedAtMs > 0 && Date.now() - startedAtMs > GAS_SYNC_TIMEOUT_MS + 15000;
  }
  if (status.state === 'queued') {
    const requestedAtMs = parseIsoTimeMs(status.requestedAt);
    return requestedAtMs > 0 && Date.now() - requestedAtMs > 30000;
  }
  return false;
}

function shouldRecoverGasSyncStatus(status) {
  if (!status?.state) return false;
  if ((status.state === 'queued' || status.state === 'running') && !gasSyncProcessing) {
    return true;
  }
  return isGasSyncStatusStale(status);
}

function scheduleGasSyncAlarm(delayMs = 50) {
  const when = Date.now() + Math.max(50, Number(delayMs) || 50);
  chrome.alarms.create(GAS_SYNC_ALARM_NAME, {
    when,
    periodInMinutes: GAS_SYNC_ALARM_PERIOD_MINUTES,
  });
}

async function getStoredGasSyncResetToken() {
  const stored = await storageGetLocal([GAS_SYNC_RESET_TOKEN_KEY]);
  const token = Number(stored[GAS_SYNC_RESET_TOKEN_KEY] || 0);
  return Number.isFinite(token) ? token : 0;
}

function clearGasSyncAlarm() {
  chrome.alarms.clear(GAS_SYNC_ALARM_NAME, () => {
    void chrome.runtime.lastError;
  });
}

async function resumeGasSyncQueueIfNeeded(delayMs = 50) {
  const stored = await storageGetLocal([GAS_SYNC_QUEUE_KEY]);
  const queue = Array.isArray(stored[GAS_SYNC_QUEUE_KEY]) ? stored[GAS_SYNC_QUEUE_KEY] : [];
  if (!queue.length) {
    clearGasSyncAlarm();
    return false;
  }
  scheduleGasSyncAlarm(delayMs);
  processGasSyncQueue().catch(() => {});
  return true;
}

async function getGasSyncStatusWithRecovery() {
  const stored = await storageGetLocal([GAS_SYNC_QUEUE_KEY, GAS_SYNC_STATUS_KEY]);
  const queue = Array.isArray(stored[GAS_SYNC_QUEUE_KEY]) ? stored[GAS_SYNC_QUEUE_KEY] : [];
  let status = stored[GAS_SYNC_STATUS_KEY] || null;

  if (!queue.length) {
    clearGasSyncAlarm();
    if (status?.state === 'queued' || status?.state === 'running') {
      gasSyncProcessing = false;
      await storageSetLocal({ [GAS_SYNC_STATUS_KEY]: null }).catch(() => {});
      return null;
    }
    return status;
  }

  if (shouldRecoverGasSyncStatus(status)) {
    gasSyncProcessing = false;
    status = buildGasSyncStatus({
      ...status,
      state: 'queued',
      startedAt: '',
      finishedAt: '',
      error: '',
      message: '⏳ 停止していたGAS書き込みを再開しています...',
    });
    await storageSetLocal({ [GAS_SYNC_STATUS_KEY]: status });
  }

  scheduleGasSyncAlarm(150);
  processGasSyncQueue().catch(() => {});
  return status;
}

// GASへGETを投げてコールドスタートを解消しておく（fire-and-forget）。
function warmupGasFromBackground(url) {
  if (!url) return;
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 15000);
  fetch(url, { method: 'GET', credentials: 'omit', redirect: 'follow', signal: controller.signal })
    .catch(() => {})
    .finally(() => clearTimeout(timeoutId));
}

async function postGasJson(url, bodyText) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), GAS_SYNC_TIMEOUT_MS);

  try {
    // GAS は POST 受信後 302 リダイレクトでレスポンスを返す。
    // Service Worker では redirect:'manual' だと Location ヘッダーが読めない
    // (opaqueredirect) ため status 0 になる。
    // redirect:'follow' なら 302→GET に変わるが、GAS 側の doPost() は
    // リダイレクト前に処理済みなので問題ない。
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: bodyText,
      credentials: 'omit',
      redirect: 'follow',
      signal: controller.signal,
    });

    return {
      status: response.status,
      responseText: await response.text(),
      responseUrl: response.url || url,
    };
  } catch (error) {
    if (error && error.name === 'AbortError') {
      throw new Error('GAS request timed out');
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

function normalizeTaiwanLookupText(value) {
  return String(value || '')
    .replace(/\u3000/g, ' ')
    .replace(/[（]/g, '(')
    .replace(/[）]/g, ')')
    .replace(/[［]/g, '[')
    .replace(/[］]/g, ']')
    .replace(/[＋]/g, '+')
    .replace(/[－–—―]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function uniqTaiwanLookupQueries(values) {
  const seen = new Set();
  const list = [];
  for (const value of values || []) {
    const text = normalizeTaiwanLookupText(value);
    const key = text.toLowerCase().replace(/[\s"'“”‘’・･:：!！?？,，、.．\-ー‐―–—~〜/／\\|()[\]{}【】「」『』<>]/g, '');
    if (!text || !key || seen.has(key)) continue;
    seen.add(key);
    list.push(text);
  }
  return list;
}

function stripTaiwanLookupNoiseForMu(value) {
  let text = normalizeTaiwanLookupText(value);
  if (!text) return '';
  const rules = [
    /\s*【[^】]*(?:首刷|初版|初回|限定|特[裝装]|通常|贈品|赠品)[^】]*】\s*/gu,
    /\s*\([^)]*(?:首刷|初版|初回|限定|特[裝装]|通常|贈品|赠品)[^)]*\)\s*/gu,
    /\s*\[[^\]]*(?:首刷|初版|初回|限定|特[裝装]|通常|贈品|赠品)[^\]]*\]\s*/gu,
    /\s*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu,
    /\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/iu,
    /\s*[#＃]?\s*[0-9０-９]{1,4}\s*$/u,
    /\s*(?:vol\.?|v\.?|第)\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|册|話|回|期|号|號)?\s*$/iu,
    /\s*(?:資料夾|资料夹|文件夾|文件夹|壓克力|压克力|立牌|色紙|色纸|貼紙|贴纸|明信片|鑰匙圈|钥匙圈|徽章|海報|海报|掛軸|挂轴|手機包|手机包|托特袋|帆布袋|手提袋|杯墊|杯垫|卡片|吊飾|吊饰|公仔|玩偶|抱枕|套組|套装|福袋|拼圖|拼图|T恤|毛巾).*/iu,
    /\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu,
  ];
  for (let i = 0; i < 5; i += 1) {
    const prev = text;
    for (const rule of rules) {
      text = text.replace(rule, ' ').trim();
    }
    text = normalizeTaiwanLookupText(text);
    if (text === prev) break;
  }
  return text;
}

function getTaiwanMuQueries(item, analysis) {
  const rawItem = item?.rawItem || item || {};
  const rowData = item?.rowData || rawItem.rowData || {};
  const rawValues = [
    analysis?.normalizedSearchTitle,
    analysis?.extractedWorkTitle,
    analysis?.originalTitle,
    rowData['原題タイトル'],
    rowData['作品名（原題）'],
    rowData['作品名原題'],
    rowData['原題商品タイトル'],
    rowData['原題商品名'],
    rawItem.原題タイトル,
    rawItem.作品名原題,
    rawItem.商品名,
    rawItem.title,
  ];
  const stripped = rawValues.map(stripTaiwanLookupNoiseForMu);
  return uniqTaiwanLookupQueries(stripped.concat(rawValues)).slice(0, 3);
}

function shouldTryTaiwanMuLookup(item, analysis) {
  const type = String(analysis?.itemType || item?.itemType || '').trim();
  if (type === 'magazine' || type === 'music_video') return false;
  return ['manga', 'bl_manga', 'goods', 'light_novel', 'novel_book', 'unknown', 'book'].includes(type) || !type;
}

// 拡張ポップアップが「追加」時に既に確定させた照会結果。
// これらが付いていれば書き込み時にMUを再照会しない（重い逐次通信の二重実行を防ぐ）。
// not_found / series_found_no_japanese も「MU側では確定」なので、再照会しても結果は変わらない。
const TAIWAN_TERMINAL_LOOKUP_STATUSES = {
  resolved: true,
  not_found: true,
  series_found_no_japanese: true,
};

async function enrichItemWithTaiwanMangaUpdates(item) {
  if (!item || typeof item !== 'object') return item;
  const existing = item.japaneseTitleLookup;
  const existingStatus = String(existing?.status || '').trim();
  if (existing && TAIWAN_TERMINAL_LOOKUP_STATUSES[existingStatus]) {
    // resolved はタイトルが入っている場合のみ信頼（空 resolved は異常なので再照会）。
    // not_found / series は「見つからない確定」なのでタイトル無しでも再照会不要。
    const hasTitle = String(existing.japaneseTitle || existing.title || '').trim();
    if (existingStatus !== 'resolved' || hasTitle) {
      return item;
    }
  }

  const rawItem = item.rawItem || item;
  const analysis = item.titleAnalysis || (
    typeof globalThis.analyzeProductTitle === 'function'
      ? globalThis.analyzeProductTitle({ ...rawItem, source: 'books_tw' })
      : null
  );
  if (!analysis || !shouldTryTaiwanMuLookup(item, analysis)) {
    return { ...item, titleAnalysis: analysis || item.titleAnalysis || null };
  }

  const queries = getTaiwanMuQueries(item, analysis);
  if (!queries.length) return { ...item, titleAnalysis: analysis };

  const client = globalThis.titleLookupMangaUpdates;
  if (!client || typeof client.lookupJapaneseTitle !== 'function') {
    return {
      ...item,
      titleAnalysis: analysis,
      japaneseTitleLookup: {
        status: 'not_found',
        provider: '',
        normalizedSearchTitle: analysis.normalizedSearchTitle || queries[0],
        extractedWorkTitle: analysis.extractedWorkTitle || queries[0],
        trace: `mangaupdatesClient:skip(no_client q=${queries[0]})`,
        errors: [],
      },
    };
  }

  try {
    const traceQuery = queries.length > 1 ? `${queries[0]}+${queries.length - 1}cand` : queries[0];
    const muResult = typeof client.lookupJapaneseTitleDetailed === 'function'
      ? await client.lookupJapaneseTitleDetailed(queries)
      : {
          status: 'resolved',
          japaneseTitle: await client.lookupJapaneseTitle(queries),
          provider: 'mangaUpdates(extension)',
          trace: '',
          candidates: [],
        };
    const japaneseTitle = normalizeTaiwanLookupText(muResult?.japaneseTitle || '');
    const status = japaneseTitle ? 'resolved' : 'not_found';
    const seriesTitles = Array.isArray(muResult?.matchedTitles)
      ? muResult.matchedTitles.map(normalizeTaiwanLookupText).filter(Boolean).slice(0, 6)
      : [];
    const detailTrace = japaneseTitle
      ? (muResult?.trace || `mangaupdatesClient:hit(q=${traceQuery})`)
      : (
          muResult?.status === 'series_found_no_japanese'
            ? `mangaupdatesClient:series_hit_no_japanese(q=${traceQuery} titles=${seriesTitles.join('/')})`
            : (muResult?.trace || `mangaupdatesClient:not_found(q=${traceQuery})`)
        );
    const normalizedTrace = detailTrace
      .replace(/^mangaUpdates:/, 'mangaupdatesClient:')
      .replace(/^aniListNative:/, 'mangaupdatesClient:aniListNative:');
    const clientTrace = normalizedTrace.indexOf('mangaupdatesClient:') === 0
      ? normalizedTrace
      : `mangaupdatesClient:${normalizedTrace}`;
    return {
      ...item,
      titleAnalysis: analysis,
      japaneseTitleLookup: {
        status,
        japaneseTitle,
        provider: japaneseTitle ? (muResult?.provider || 'mangaUpdates(extension)') : '',
        normalizedSearchTitle: analysis.normalizedSearchTitle || queries[0],
        extractedWorkTitle: analysis.extractedWorkTitle || queries[0],
        score: japaneseTitle ? 900 : 0,
        trace: clientTrace,
        candidates: japaneseTitle
          ? [{ title: japaneseTitle, provider: muResult?.provider || 'mangaUpdates(extension)', score: 900 }]
          : (Array.isArray(muResult?.candidates) ? muResult.candidates : []),
        errors: Array.isArray(muResult?.errors) ? muResult.errors : [],
      },
    };
  } catch (error) {
    return {
      ...item,
      titleAnalysis: analysis,
      japaneseTitleLookup: {
        status: 'partial_error',
        provider: 'mangaUpdates(extension)',
        normalizedSearchTitle: analysis.normalizedSearchTitle || queries[0],
        extractedWorkTitle: analysis.extractedWorkTitle || queries[0],
        trace: `mangaupdatesClient:error(q=${queries[0]} ${error?.message || error})`,
        errors: [String(error?.message || error || 'MangaUpdates lookup failed')],
      },
    };
  }
}

async function enrichItemsWithTaiwanMangaUpdates(items) {
  const sourceItems = Array.isArray(items) ? items : [];
  const enriched = [];
  for (const item of sourceItems) {
    let timeoutId;
    const lookupPromise = enrichItemWithTaiwanMangaUpdates(item).then(res => {
      if (timeoutId) clearTimeout(timeoutId);
      return res;
    });
    const timeoutPromise = new Promise(resolve => {
      timeoutId = setTimeout(() => {
        const rawItem = item?.rawItem || item || {};
        const analysis = item?.titleAnalysis || (
          typeof globalThis.analyzeProductTitle === 'function'
            ? globalThis.analyzeProductTitle({ ...rawItem, source: 'books_tw' })
            : null
        );
        resolve(Object.assign({}, item, {
          titleAnalysis: analysis,
          japaneseTitleLookup: item.japaneseTitleLookup || {
            status: 'skipped',
            provider: 'mangaUpdates(extension)',
            normalizedSearchTitle: analysis?.normalizedSearchTitle || '',
            extractedWorkTitle: analysis?.extractedWorkTitle || '',
            trace: 'mangaupdatesClient:skipped(timeout_during_write)',
            errors: ['MangaUpdates lookup timed out']
          }
        }));
      }, 3000);
    });
    enriched.push(await Promise.race([lookupPromise, timeoutPromise]));
  }
  return enriched;
}

async function postProductsToGas(url, items) {
  const enrichedItems = await enrichItemsWithTaiwanMangaUpdates(items);
  const response = await postGasJson(url, JSON.stringify({
    action: 'upsertProductWithLookup',
    source: 'books_tw',
    appendMode: 'append',
    requestedAt: new Date().toISOString(),
    items: enrichedItems,
  }));

  const responseText = response.responseText;
  let data = null;
  try {
    data = responseText ? JSON.parse(responseText) : null;
  } catch {
    data = null;
  }

  if (response.status < 200 || response.status >= 300) {
    throw new Error(data?.error || `HTTP ${response.status}`);
  }

  if (!data || typeof data !== 'object') {
    const snippet = String(responseText || '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 160);
    throw new Error(`GAS JSON応答を確認できませんでした: ${snippet || 'empty response'}`);
  }

  if (data.ok === false) {
    throw new Error(data.error || 'GAS write failed');
  }

  return data;
}

function normalizeGasWriteErrorMessage(errorLike) {
  const raw = String(errorLike?.message || errorLike || '').trim();
  if (!raw) return 'GAS write failed';
  const compact = raw.replace(/\s+/g, ' ').trim();
  const build = (title, details = [], actions = []) => {
    const body = [title];
    if (details.length) body.push(`原因候補: ${details.join(' / ')}`);
    if (actions.length) body.push(`対処: ${actions.join(' / ')}`);
    body.push(`原文: ${compact}`);
    return body.join('\n');
  };

  if (/データの行数が範囲の行数と一致しません/.test(raw)) {
    return build(
      '行数不一致エラー（setValues）',
      ['重複判定の途中で書込配列が崩れた', 'シート列構成と送信データ構成の不一致'],
      ['同一商品コード/URL の重複行を確認', 'GASヘッダー列と拡張列定義の差分を確認', '対象行を1件に絞って再送']
    );
  }
  if (/このセルで設定しているデータの入力規則に違反しています/.test(raw)) {
    return build(
      '入力規則エラー（ドロップダウン/検証ルール違反）',
      ['送信値が許可候補にない', 'マスター未登録の値が送られた'],
      ['対象列の入力規則候補に値を追加', '未確定値は空欄送信して備考へ退避']
    );
  }
  if (/範囲が見つかりません|指定した範囲|対象のシートが見つかりません|does not exist|Cannot find range/i.test(raw)) {
    return build(
      'シート/範囲エラー',
      ['対象シート名が変更された', 'ヘッダー行や列位置が想定とずれている'],
      ['GAS対象シート名設定を確認', 'ヘッダー行を再生成し列名一致を確認']
    );
  }
  if (/HTTP 401|HTTP 403|権限|アクセス権|permission|forbidden|unauthorized/i.test(raw)) {
    return build(
      '認証/権限エラー',
      ['Webアプリ実行ユーザーの権限不足', 'GAS公開範囲が不一致'],
      ['GASを再デプロイして権限見直し', '拡張のGAS URLが最新か確認']
    );
  }
  if (/timed out|timeout|hard timeout|ネットワーク|network|fetch failed/i.test(raw)) {
    return build(
      '通信タイムアウト/ネットワークエラー',
      ['一時的な通信失敗', 'GAS側処理時間超過'],
      ['少し待って再実行', '送信件数を減らして分割実行']
    );
  }
  return build('GAS書き込みエラー');
}

function createGasSyncJob(request) {
  return {
    jobId: `gas_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    url: String(request.url || '').trim(),
    items: Array.isArray(request.items) ? request.items : [],
    requestedAt: new Date().toISOString(),
    attempts: 0,
  };
}

// タイムアウト/ネットワーク系の一過性エラーのみ自動リトライ対象とする。
// 行数不一致・入力規則違反・権限エラー等のロジック系は再送しても無駄なので除外。
function isRetryableGasSyncError(errorLike) {
  const raw = String(errorLike?.message || errorLike || '');
  return /timed out|timeout|hard timeout|ネットワーク|network|fetch failed|failed to fetch|load failed|connection|ECONN|socket/i.test(raw);
}

function delayMs_(ms) {
  return new Promise(resolve => setTimeout(resolve, Math.max(0, Number(ms) || 0)));
}

async function queueGasSyncJob(request) {
  const job = createGasSyncJob(request);
  if (!job.url) throw new Error('GAS URL が未設定です');
  if (!job.items.length) throw new Error('送信対象の商品がありません');

  const stored = await storageGetLocal([GAS_SYNC_QUEUE_KEY]);
  const queue = Array.isArray(stored[GAS_SYNC_QUEUE_KEY]) ? stored[GAS_SYNC_QUEUE_KEY] : [];
  queue.push(job);

  const status = buildGasSyncStatus({
    state: 'queued',
    jobId: job.jobId,
    itemCount: job.items.length,
    queueLength: queue.length,
    requestedAt: job.requestedAt,
    message: `⏳ ${job.items.length}件のシート書き込みをバックグラウンド送信待ちにしました`,
  });

  await storageSetLocal({
    [GAS_SYNC_QUEUE_KEY]: queue,
    [GAS_SYNC_STATUS_KEY]: status,
  });

  scheduleGasSyncAlarm(50);
  processGasSyncQueue().catch(() => {});
  return status;
}

async function processGasSyncQueue() {
  if (gasSyncProcessing) return;
  gasSyncProcessing = true;

  try {
    while (true) {
      const stored = await storageGetLocal([GAS_SYNC_QUEUE_KEY]);
      const queue = Array.isArray(stored[GAS_SYNC_QUEUE_KEY]) ? stored[GAS_SYNC_QUEUE_KEY] : [];
      if (!queue.length) {
        clearGasSyncAlarm();
        return;
      }

      const job = queue[0];
      const nextQueue = queue.slice(1);
      const runToken = await getStoredGasSyncResetToken();
      gasSyncResetToken = runToken;

      await storageSetLocal({
        [GAS_SYNC_STATUS_KEY]: buildGasSyncStatus({
          state: 'running',
          jobId: job.jobId,
          itemCount: job.items.length,
          queueLength: queue.length,
          requestedAt: job.requestedAt,
          startedAt: new Date().toISOString(),
          message: `📝 ${job.items.length}件をGASへ書き込み中...`,
        }),
      });

      try {
        const result = await postProductsToGas(job.url, job.items);
        const latestResetToken = await getStoredGasSyncResetToken();
        gasSyncResetToken = latestResetToken;
        if (runToken !== latestResetToken) continue;
        const resultList = Array.isArray(result?.results) ? result.results : [];
        const appended = Number.isFinite(Number(result?.appended))
          ? Number(result.appended)
          : resultList.filter(entry => entry && entry.ok && !entry.skipped).length;
        const skipped = Number.isFinite(Number(result?.skipped))
          ? Number(result.skipped)
          : resultList.filter(entry => entry && entry.skipped).length;
        const firstAppendedResult = resultList.find(entry => entry && entry.ok && !entry.skipped) || null;
        const locationText = appended === 1 && firstAppendedResult?.sheetName && firstAppendedResult?.appendedRow
          ? `（${firstAppendedResult.sheetName} ${firstAppendedResult.appendedRow}行目）`
          : '';
        const message = appended > 0
          ? `✅ シートへ ${appended} 件追記しました${skipped > 0 ? `（重複スキップ:${skipped}）` : ''}${locationText}`
          : skipped > 0
            ? `ℹ️ 重複のため新規追加はありませんでした（スキップ:${skipped}）`
            : 'ℹ️ 追加対象はありませんでした';

        await storageSetLocal({
          [GAS_SYNC_QUEUE_KEY]: nextQueue,
          [GAS_SYNC_STATUS_KEY]: buildGasSyncStatus({
            state: 'success',
            jobId: job.jobId,
            itemCount: job.items.length,
            queueLength: nextQueue.length,
            requestedAt: job.requestedAt,
            finishedAt: new Date().toISOString(),
            appended,
            skipped,
            result,
            message,
          }),
        });
      } catch (error) {
        const latestResetToken = await getStoredGasSyncResetToken();
        gasSyncResetToken = latestResetToken;
        if (runToken !== latestResetToken) continue;

        const attemptsSoFar = Number(job.attempts || 0) + 1;
        // 一過性（タイムアウト/通信失敗）かつ上限未満なら、同一ジョブを先頭に戻して自動再送する。
        if (isRetryableGasSyncError(error) && attemptsSoFar < GAS_SYNC_MAX_ATTEMPTS) {
          const retryJob = { ...job, attempts: attemptsSoFar };
          const backoffMs = GAS_SYNC_RETRY_BASE_DELAY_MS * Math.pow(2, attemptsSoFar - 1);
          await storageSetLocal({
            [GAS_SYNC_QUEUE_KEY]: [retryJob].concat(nextQueue),
            [GAS_SYNC_STATUS_KEY]: buildGasSyncStatus({
              state: 'queued',
              jobId: job.jobId,
              itemCount: job.items.length,
              queueLength: nextQueue.length + 1,
              // バックオフ待機中に「stale（30秒超の queued）」と誤判定され
              // 二重処理されないよう、requestedAt は現在時刻に更新する。
              requestedAt: new Date().toISOString(),
              message: `⏳ 通信が不安定なため再送します（${attemptsSoFar}/${GAS_SYNC_MAX_ATTEMPTS - 1}回目・${Math.round(backoffMs / 1000)}秒後）`,
            }),
          });
          // GASのコールドスタートを温め直してから待機・再試行する。
          try { warmupGasFromBackground(job.url); } catch (_) {}
          await delayMs_(backoffMs);
          continue;
        }

        const normalizedError = normalizeGasWriteErrorMessage(error);
        const attemptSuffix = attemptsSoFar > 1 ? `（${attemptsSoFar}回試行）` : '';
        await storageSetLocal({
          [GAS_SYNC_QUEUE_KEY]: nextQueue,
          [GAS_SYNC_STATUS_KEY]: buildGasSyncStatus({
            state: 'error',
            jobId: job.jobId,
            itemCount: job.items.length,
            queueLength: nextQueue.length,
            requestedAt: job.requestedAt,
            finishedAt: new Date().toISOString(),
            error: normalizedError,
            message: `❌ ${normalizedError}${attemptSuffix}`,
          }),
        });
      }
    }
  } finally {
    gasSyncProcessing = false;
  }
}

async function resetGasSyncState(request = {}) {
  const requestedToken = Number(request?.resetToken || 0);
  const currentToken = await getStoredGasSyncResetToken();
  const nextToken = requestedToken > 0 ? requestedToken : currentToken + 1;
  gasSyncResetToken = nextToken;
  gasSyncProcessing = false;
  clearGasSyncAlarm();
  await storageSetLocal({
    [GAS_SYNC_QUEUE_KEY]: [],
    [GAS_SYNC_STATUS_KEY]: null,
    [GAS_SYNC_RESET_TOKEN_KEY]: nextToken,
  });
}

function sanitizePathSegment(segment) {
  return String(segment || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
}

function inferFilenameFromUrl(url, fallbackName) {
  try {
    if (/^https?:\/\//i.test(String(url || ''))) {
      const pathname = new URL(url).pathname || '';
      const basename = sanitizePathSegment(pathname.split('/').pop() || '');
      if (basename) return basename;
    }
  } catch {
    // ignore URL parse errors
  }
  return sanitizePathSegment(fallbackName || 'download.bin') || 'download.bin';
}

function sanitizeRelativePath(pathText, fallbackName) {
  const normalized = String(pathText || '').replace(/\\/g, '/');
  const parts = normalized
    .split('/')
    .map(sanitizePathSegment)
    .filter(Boolean);

  if (parts.length === 0) {
    return sanitizePathSegment(fallbackName || 'download.bin') || 'download.bin';
  }

  return parts.join('/');
}

function bytesToBase64(bytes) {
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}




function inferTaiwanProductCodeFromFilename(filename = '') {
  const firstSegment = String(filename || '')
    .replace(/\\/g, '/')
    .split('/')
    .map(part => String(part || '').trim())
    .find(Boolean) || '';
  return sanitizePathSegment(firstSegment);
}



function getPendingTaiwanFilenameValue(entry) {
  if (!entry) return '';
  if (typeof entry === 'string') return entry;
  return String(entry.filename || '').trim();
}

function prunePendingTaiwanFilenameMarkers(now = Date.now()) {
  for (const [markerFilename, entry] of pendingTaiwanFilenameByMarker.entries()) {
    const createdAt = Number(entry?.createdAt || 0);
    if (createdAt && now - createdAt > TAIWAN_FILENAME_HOOK_TTL_MS) {
      pendingTaiwanFilenameByMarker.delete(markerFilename);
    }
  }
}

function releaseTaiwanFilenameHookIfIdle() {
  prunePendingTaiwanFilenameMarkers();
  if (pendingTaiwanFilenameByMarker.size > 0) return;
  if (taiwanFilenameHookCleanupTimer) {
    clearTimeout(taiwanFilenameHookCleanupTimer);
    taiwanFilenameHookCleanupTimer = null;
  }
  if (taiwanFilenameHookRegistered) {
    chrome.downloads.onDeterminingFilename.removeListener(handleTaiwanDeterminingFilename);
    taiwanFilenameHookRegistered = false;
  }
}

function scheduleTaiwanFilenameHookCleanup() {
  if (taiwanFilenameHookCleanupTimer) {
    clearTimeout(taiwanFilenameHookCleanupTimer);
  }
  taiwanFilenameHookCleanupTimer = setTimeout(() => {
    releaseTaiwanFilenameHookIfIdle();
  }, TAIWAN_FILENAME_HOOK_TTL_MS);
}

function ensureTaiwanFilenameHook() {
  if (!taiwanFilenameHookRegistered) {
    chrome.downloads.onDeterminingFilename.addListener(handleTaiwanDeterminingFilename);
    taiwanFilenameHookRegistered = true;
  }
  scheduleTaiwanFilenameHookCleanup();
}

function buildTaiwanMarkerFilename(targetFilename = 'download.bin') {
  const safeTargetFilename = sanitizeRelativePath(targetFilename, 'download.bin');
  const extensionMatch = safeTargetFilename.match(/(\.[a-z0-9]+)$/i);
  const extension = extensionMatch ? extensionMatch[1] : '.bin';
  const token = `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const markerFilename = `${TAIWAN_DOWNLOAD_MARKER}${token}${extension}`;
  pendingTaiwanFilenameByMarker.set(markerFilename, {
    filename: safeTargetFilename,
    createdAt: Date.now(),
  });
  ensureTaiwanFilenameHook();
  return markerFilename;
}

function consumePendingTaiwanFilenameByMarker(itemFilename = '') {
  const normalized = String(itemFilename || '')
    .replace(/\\/g, '/')
    .split('/')
    .map(part => String(part || '').trim())
    .filter(Boolean)
    .pop() || '';
  if (!normalized || !normalized.includes(TAIWAN_DOWNLOAD_MARKER)) return '';

  const direct = pendingTaiwanFilenameByMarker.get(normalized);
  if (direct) {
    pendingTaiwanFilenameByMarker.delete(normalized);
    return getPendingTaiwanFilenameValue(direct);
  }

  for (const [markerFilename, entry] of pendingTaiwanFilenameByMarker.entries()) {
    if (normalized.includes(markerFilename)) {
      pendingTaiwanFilenameByMarker.delete(markerFilename);
      return getPendingTaiwanFilenameValue(entry);
    }
  }

  return '';
}function isGoodsLikeTaiwanProductCode(productCode = '') {
  return /^[NM]/i.test(String(productCode || '').trim());
}

function resolveAbsoluteUrl(url, baseUrl = 'https://www.books.com.tw/') {
  try {
    return new URL(String(url || '').trim(), baseUrl).href;
  } catch {
    return String(url || '').trim();
  }
}

function buildTaiwanDirectImageVariants(urlText, options = {}) {
  const keepThumbVariant = options.keepThumbVariant === true;
  const raw = resolveAbsoluteUrl(String(urlText || '').trim());
  if (!/^https?:\/\//i.test(raw)) return [];

  const variants = [];
  const pushVariant = candidate => {
    const normalized = resolveAbsoluteUrl(String(candidate || '').trim());
    if (!/^https?:\/\//i.test(normalized)) return;
    if (variants.includes(normalized)) return;
    variants.push(normalized);
  };
  const pushExtVariants = candidate => {
    const text = String(candidate || '').trim();
    if (!text) return;
    if (/\.webp(?=($|[?#]))/i.test(text)) {
      pushVariant(text.replace(/\.webp(?=($|[?#]))/i, '.jpeg'));
      pushVariant(text.replace(/\.webp(?=($|[?#]))/i, '.jpg'));
    }
    pushVariant(text);
  };

  const preferredSlot = raw.replace(
    /_t_(\d+)(?=\.(?:jpe?g|png|webp)(?:$|[?#]))/gi,
    keepThumbVariant ? '_t_$1' : '_b_$1'
  );

  pushExtVariants(preferredSlot);
  if (preferredSlot !== raw) {
    pushExtVariants(raw);
  }
  return variants;
}

function upgradeTaiwanDirectImageVariant(urlText, options = {}) {
  return buildTaiwanDirectImageVariants(urlText, options)[0] || resolveAbsoluteUrl(urlText);
}

function buildTaiwanBannerWrapperUrl(imageUrl) {
  const resolved = resolveAbsoluteUrl(String(imageUrl || '').trim());
  if (!/^https?:\/\/addons\.books\.com\.tw\/G\/ADbanner\//i.test(resolved)) return '';
  return `https://im1.book.com.tw/image/getImage?i=${encodeURIComponent(resolved)}`;
}

function unwrapTaiwanImageSource(urlText, options = {}) {
  const raw = String(urlText || '').trim();
  if (!raw) return '';

  try {
    const parsed = new URL(raw, 'https://www.books.com.tw/');
    const original = parsed.searchParams.get('i');
    if (original) {
      let originalUrl = decodeURIComponent(original);
      if (originalUrl.startsWith('//')) originalUrl = `https:${originalUrl}`;
      const wrappedBannerUrl = buildTaiwanBannerWrapperUrl(originalUrl);
      if (wrappedBannerUrl) return wrappedBannerUrl;
      return upgradeTaiwanDirectImageVariant(originalUrl, options);
    }
    const upgradedUrl = upgradeTaiwanDirectImageVariant(parsed.href, options);
    return buildTaiwanBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  } catch {
    const upgradedUrl = upgradeTaiwanDirectImageVariant(raw, options);
    return buildTaiwanBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  }
}

function buildTaiwanImageCandidates(url, options = {}) {
  const source = unwrapTaiwanImageSource(url, options);
  const candidates = [];
  const pushCandidate = raw => {
    const normalized = resolveAbsoluteUrl(String(raw || '').trim());
    if (!/^https?:\/\//i.test(normalized)) return;
    if (candidates.includes(normalized)) return;
    candidates.push(normalized);
  };

  pushCandidate(source);
  buildTaiwanDirectImageVariants(source, options).forEach(pushCandidate);

  const rawText = String(url || '').trim();
  if (rawText && !/\/image\/getImage\?|\/fancybox\/getImage\.php\?/i.test(rawText)) {
    buildTaiwanDirectImageVariants(rawText, options).forEach(pushCandidate);
  }
  return candidates;
}

function extractImageCandidateFromHtml(htmlText, baseUrl) {
  const html = String(htmlText || '');
  const candidates = [];
  const pushCandidate = raw => {
    const normalized = resolveAbsoluteUrl(raw, baseUrl);
    if (!/^https?:\/\//i.test(normalized)) return;
    candidates.push(normalized);
  };

  Array.from(html.matchAll(/https?:\/\/[^"'\s<)]+(?:jpe?g|png|webp)(?:\?[^"'\s<)]*)?/gi)).forEach(match => pushCandidate(match[0]));
  Array.from(html.matchAll(/https?:\/\/[^"'\s<)]*(?:image\/getImage\?|fancybox\/getImage\.php\?)[^"'\s<)]*/gi)).forEach(match => pushCandidate(match[0]));
  Array.from(html.matchAll(/(?:src|href)=["']([^"']+(?:jpe?g|png|webp)(?:\?[^"']*)?)["']/gi)).forEach(match => pushCandidate(match[1]));
  Array.from(html.matchAll(/(?:src|href)=["']([^"']*(?:image\/getImage\?|fancybox\/getImage\.php\?)[^"']*)["']/gi)).forEach(match => pushCandidate(match[1]));

  return candidates.find(Boolean) || '';
}

async function fetchTaiwanImageBlob(url, depth = 0) {
  const response = await fetch(url, {
    credentials: 'include',
    redirect: 'follow',
    cache: 'no-store',
  });
  if (!response.ok) {
    throw new Error(`image fetch failed: HTTP ${response.status}`);
  }

  const contentType = String(response.headers.get('content-type') || '').toLowerCase();
  if (contentType.startsWith('image/')) {
    return await response.blob();
  }

  const html = await response.text();
  if (depth >= 1) {
    throw new Error(`image fetch returned ${contentType || 'non-image response'}`);
  }

  const nestedUrl = extractImageCandidateFromHtml(html, response.url || url);
  if (nestedUrl && nestedUrl !== url) {
    return fetchTaiwanImageBlob(nestedUrl, depth + 1);
  }

  throw new Error(`image fetch returned ${contentType || 'non-image response'}`);
}

async function convertBlobToJpegDataUrl(blob, maxEdge = TAIWAN_IMAGE_MAX_EDGE) {
  if (!blob) throw new Error('image blob missing');
  if (typeof createImageBitmap !== 'function' || typeof OffscreenCanvas === 'undefined') {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    return `data:${blob.type || 'image/jpeg'};base64,${bytesToBase64(bytes)}`;
  }

  const bitmap = await createImageBitmap(blob);
  try {
    const longestEdge = Math.max(bitmap.width || 0, bitmap.height || 0, 1);
    const scale = Math.min(1, maxEdge / longestEdge);
    const width = Math.max(1, Math.round((bitmap.width || 1) * scale));
    const height = Math.max(1, Math.round((bitmap.height || 1) * scale));
    const canvas = new OffscreenCanvas(width, height);
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('canvas context unavailable');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(bitmap, 0, 0, width, height);
    const jpegBlob = await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.92 });
    const bytes = new Uint8Array(await jpegBlob.arrayBuffer());
    return `data:image/jpeg;base64,${bytesToBase64(bytes)}`;
  } finally {
    if (typeof bitmap.close === 'function') bitmap.close();
  }
}

function buildTaiwanImageDownloadUrl(url, options = {}) {
  const text = String(url || '').trim();
  if (!/^https?:\/\//i.test(text)) return text;
  const candidates = buildTaiwanImageCandidates(text, options);
  return candidates.find(candidate => /^https?:\/\//i.test(candidate)) || text;
}


function handleTaiwanDeterminingFilename(item, suggest) {
  prunePendingTaiwanFilenameMarkers();
  const desiredFilename = consumePendingTaiwanFilenameByMarker(item?.filename || '');
  if (desiredFilename) {
    suggest({ filename: desiredFilename, conflictAction: 'uniquify' });
  } else {
    suggest();
  }
  releaseTaiwanFilenameHookIfIdle();
}
chrome.alarms.onAlarm.addListener(alarm => {
  if (alarm?.name === GAS_SYNC_ALARM_NAME) {
    processGasSyncQueue().catch(() => {});
  }
});

chrome.runtime.onStartup.addListener(() => {
  resumeGasSyncQueueIfNeeded(200).catch(() => {});
});

chrome.runtime.onInstalled.addListener(() => {
  resumeGasSyncQueueIfNeeded(200).catch(() => {});
});

resumeGasSyncQueueIfNeeded(250).catch(() => {});

function startDownload(request, sendResponse) {
  const { url, filename, saveAs } = request || {};
  if (!url) {
    sendResponse({ ok: false, error: 'url missing' });
    return;
  }

  const planned = sanitizeRelativePath(filename, inferFilenameFromUrl(url, 'download.bin'));
  chrome.downloads.download(
    {
      url,
      filename: planned,
      saveAs: saveAs === true,
      conflictAction: 'uniquify',
    },
    downloadId => {
      if (chrome.runtime.lastError || !downloadId) {
        sendResponse({ ok: false, error: chrome.runtime.lastError?.message || 'download failed' });
        return;
      }
      sendResponse({ ok: true, downloadId, filename: planned });
    }
  );
}

async function prepareTaiwanImageDownload(request) {
  const url = String(request.url || '').trim();
  const filename = sanitizeRelativePath(request.filename, '1.jpg');
  if (!url) {
    throw new Error('url missing');
  }

  const productCode = inferTaiwanProductCodeFromFilename(filename);
  const sequence = Number(request.sequence) || 0;
  const isMainBookImage = sequence === 1 && !isGoodsLikeTaiwanProductCode(productCode);
  const options = {
    keepThumbVariant: request.keepThumbVariant === true
      || isGoodsLikeTaiwanProductCode(productCode)
      || isMainBookImage,
  };
  const candidates = buildTaiwanImageCandidates(url, options);
  const tryUrls = candidates.length ? candidates : [buildTaiwanImageDownloadUrl(url, options)];

  let blob = null;
  let lastError = '';
  for (const candidate of tryUrls) {
    try {
      blob = await fetchTaiwanImageBlob(candidate);
      if (blob) break;
    } catch (error) {
      lastError = String(error?.message || error || 'image fetch failed');
    }
  }
  if (!blob) {
    throw new Error(`image prepare failed: ${lastError || 'all candidates failed'}`);
  }

  const maxEdge = isMainBookImage ? 800 : TAIWAN_IMAGE_MAX_EDGE;
  const dataUrl = await convertBlobToJpegDataUrl(blob, maxEdge);
  const downloadName = buildTaiwanMarkerFilename(filename);
  return { ok: true, url: dataUrl, downloadName, filename };
}

async function prepareTaiwanCsvDownload(request) {
  const csvText = String(request.csvText || '');
  const product = request.product && typeof request.product === 'object' ? request.product : null;
  const fallbackFilename = product
    ? `${getDownloadCode(product)}.csv`
    : 'export.csv';
  const filename = sanitizeRelativePath(request.filename, fallbackFilename);
  if (!csvText) {
    throw new Error('csv text missing');
  }

  const csvBytes = new TextEncoder().encode(csvText);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${bytesToBase64(csvBytes)}`;
  const downloadName = buildTaiwanMarkerFilename(filename);
  return { ok: true, url: dataUrl, downloadName, filename };
}

function handleImageDownload(request, sendResponse) {
  const url = String(request.url || '').trim();
  const filename = sanitizeRelativePath(request.filename, '1.jpg');
  if (!url) {
    sendResponse({ ok: false, error: 'url missing' });
    return;
  }

  const saveAs = request.saveAs === true;
  const productCode = inferTaiwanProductCodeFromFilename(filename);
  const preparedUrl = buildTaiwanImageDownloadUrl(url, {
    keepThumbVariant: isGoodsLikeTaiwanProductCode(productCode),
  });
  startDownload({ url: preparedUrl, filename, saveAs }, sendResponse);
}

function handleCsvDownload(request, sendResponse) {
  const csvText = String(request.csvText || '');
  const product = request.product && typeof request.product === 'object' ? request.product : null;
  const fallbackFilename = product
    ? `${getDownloadCode(product)}.csv`
    : 'export.csv';
  const filename = sanitizeRelativePath(request.filename, fallbackFilename);
  if (!csvText) {
    sendResponse({ ok: false, error: 'csv text missing' });
    return;
  }

  const csvBytes = new TextEncoder().encode(csvText);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${bytesToBase64(csvBytes)}`;
  startDownload({ url: dataUrl, filename, saveAs: request.saveAs === true }, sendResponse);
}

function downloadImageTask(url, filename, saveAs = false, extra = {}) {
  return new Promise(resolve => {
    handleImageDownload({ url, filename, saveAs, ...extra }, response => {
      resolve(response?.ok ? { ok: true } : { ok: false, error: response?.error || 'download failed' });
    });
  });
}

function downloadCsvTask(csvText, filename, saveAs = false, extra = {}) {
  return new Promise(resolve => {
    handleCsvDownload({ csvText, filename, saveAs, ...extra }, response => {
      resolve(Boolean(response?.ok));
    });
  });
}

function extractProductCode(product) {
  const candidates = [
    product?.商品コード,
    product?.siteProductCode,
    product?.博客來商品コード,
    product?.productCode,
    product?.ISBN,
    product?.isbn,
  ];

  for (const value of candidates) {
    const directCode = String(value || '').trim();
    if (directCode) return directCode;
  }

  const pageUrl = String(product?.URL || product?.url || '').trim();
  if (!pageUrl) return '';

  try {
    const parsed = new URL(pageUrl);
    const byPath = parsed.pathname.match(/\/products\/([^/?#]+)/i)?.[1] || '';
    if (byPath) return byPath.trim();
  } catch {
    const byText = pageUrl.match(/\/products\/([^/?#]+)/i)?.[1] || '';
    if (byText) return byText.trim();
  }

  return '';
}

function getDownloadCode(product) {
  const code = extractProductCode(product);
  if (code) return sanitizePathSegment(code) || 'item';
  return sanitizePathSegment(product?.商品名 || product?.ページタイトル || product?.title || 'item') || 'item';
}

function guessExt(url) {
  return IMAGE_FILENAME_EXT;
}

function normalizeAdditionalImages(product) {
  return String(
    product?.追加画像URL || [
      product?.画像URL_2枚目 || '',
      product?.画像URL_3枚目 || '',
      product?.画像URL_4枚目以降 || '',
    ].filter(Boolean).join(';')
  )
    .replace(/\r?\n+/g, ';')
    .replace(/;{2,}/g, ';')
    .replace(/^;|;$/g, '');
}

function collectProductImageUrls(product) {
  const ordered = [];
  const seen = new Set();

  const pushUrl = value => {
    const url = String(value || '').trim();
    if (!/^https?:\/\//i.test(url)) return;
    if (seen.has(url)) return;
    seen.add(url);
    ordered.push(url);
  };

  pushUrl(product?.画像URL || product?.imageUrl || '');
  normalizeAdditionalImages(product)
    .split(';')
    .map(url => url.trim())
    .filter(Boolean)
    .forEach(pushUrl);

  return ordered;
}

function buildCsvContent(product) {
  const headers = [
    '商品コード',
    '商品名',
    '価格',
    '発売日',
    'メイン画像URL',
    '追加画像URL',
    '商品ページURL',
    '商品説明',
    'ISBN',
    '著者',
    '翻訳者',
    'イラストレーター',
    '出版社',
  ];

  const row = [
    extractProductCode(product),
    product?.商品名 || '',
    product?.価格 || '',
    product?.発売日 || '',
    product?.画像URL || '',
    normalizeAdditionalImages(product),
    product?.URL || '',
    product?.商品説明 || '',
    product?.ISBN || '',
    product?.著者 || '',
    product?.翻訳者 || '',
    product?.イラストレーター || '',
    product?.出版社 || '',
  ];

  return '\uFEFF' + [headers, row]
    .map(record => record.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
    .join('\r\n');
}

async function processSaveProductsAssetsJob(job) {
  const products = Array.isArray(job.products) ? job.products : [];
  const mode = job.mode === 'images_only' ? 'images_only' : 'csv_and_images';
  const saveAs = job.saveAs !== false;
  if (!products.length) {
    throw new Error('保存対象の商品がありません');
  }

  let totalImages = 0;
  let successImages = 0;
  let csvCount = 0;

  for (const product of products) {
    const downloadCode = getDownloadCode(product);
    const rootedFolder = downloadCode;

    if (mode === 'csv_and_images') {
      const csvContent = buildCsvContent(product);
      const csvOk = await downloadCsvTask(csvContent, `${downloadCode}.csv`, saveAs, { product });
      if (csvOk) csvCount += 1;
    }

    const imageUrls = collectProductImageUrls(product);
    let sequence = 1;
    for (const url of imageUrls) {
      totalImages += 1;
      const ext = guessExt(url);
      const result = await downloadImageTask(
        url,
        `${rootedFolder}/${sequence}.${ext}`,
        saveAs,
        { product, sequence }
      );
      if (result.ok) successImages += 1;
      sequence += 1;
    }
  }

  return {
    mode,
    productCount: products.length,
    totalImages,
    successImages,
    failedImages: Math.max(0, totalImages - successImages),
    csvCount,
  };
}
function enqueueDownloadJobs(request, sendResponse) {
  const jobs = Array.isArray(request.jobs) ? request.jobs : [];
  if (!jobs.length) {
    sendResponse({ ok: false, error: 'jobs missing', accepted: 0, total: 0 });
    return;
  }

  // nextExpectedFilename が上書きされないよう、1件ずつ逐次実行する
  (async () => {
    let accepted = 0;
    for (const job of jobs) {
      if (!job || typeof job !== 'object') continue;

      const kind = String(job.kind || '').toLowerCase();
      if (kind === 'image') {
        const url = String(job.url || '').trim();
        const filename = sanitizeRelativePath(job.filename, '1.jpg');
        if (!url) continue;

        await new Promise(resolve => {
          const preparedUrl = buildTaiwanImageDownloadUrl(url);
          startDownload({ url: preparedUrl, filename, saveAs: false }, () => resolve());
        });
      }

      if (kind === 'csv') {
        const csvText = String(job.csvText || '');
        const filename = sanitizeRelativePath(job.filename, 'export.csv');
        if (!csvText) continue;

        const csvBytes = new TextEncoder().encode(csvText);
        const dataUrl = `data:text/csv;charset=utf-8;base64,${bytesToBase64(csvBytes)}`;
        await new Promise(resolve => {
          startDownload({ url: dataUrl, filename, saveAs: false }, () => resolve());
        });
        accepted += 1;
      }
    }

    sendResponse({ ok: true, accepted, total: jobs.length });
  })();
}

chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request.action === 'queueGasPost') {
    queueGasSyncJob(request)
      .then(status => sendResponse({ ok: true, status }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'queue failed' }));
    return true;
  }

  if (request.action === 'getGasSyncStatus') {
    getGasSyncStatusWithRecovery()
      .then(status => sendResponse({ ok: true, status }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'status read failed' }));
    return true;
  }

  if (request.action === 'clearGasSyncStatus') {
    storageSetLocal({ [GAS_SYNC_STATUS_KEY]: null })
      .then(() => sendResponse({ ok: true }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'status clear failed' }));
    return true;
  }

  if (request.action === 'resetGasSyncState') {
    resetGasSyncState(request)
      .then(() => sendResponse({ ok: true }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'reset failed' }));
    return true;
  }

  if (request.action === 'lookupMangaUpdatesJapaneseTitleDirect') {
    const client = globalThis.titleLookupMangaUpdates;
    const queries = Array.isArray(request.queries) ? request.queries : [];
    if (!client || typeof client.lookupJapaneseTitleDetailed !== 'function') {
      sendResponse({ ok: false, error: 'MangaUpdates client unavailable' });
      return true;
    }
    client.lookupJapaneseTitleDetailed(queries)
      .then(result => sendResponse({ ok: true, result }))
      .catch(error => sendResponse({ ok: false, error: error?.message || 'lookup failed' }));
    return true;
  }

  if (request.action === 'prepareTaiwanImageDownload') {
    prepareTaiwanImageDownload(request)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ ok: false, error: error.message || 'prepare image failed' }));
    return true;
  }

  if (request.action === 'prepareTaiwanCsvDownload') {
    prepareTaiwanCsvDownload(request)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ ok: false, error: error.message || 'prepare csv failed' }));
    return true;
  }

  if (request.action === 'taiwanDownloadImage' || request.action === '__legacy_taiwan_downloadImage') {
    handleImageDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'taiwanDownloadCsv' || request.action === '__legacy_taiwan_downloadCsv') {
    handleCsvDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'enqueueDownloads') {
    enqueueDownloadJobs(request, sendResponse);
    return true;
  }

  if (request.action === 'enqueueTaiwanSaveProductsAssets' || request.action === '__legacy_taiwan_saveProductsAssets') {
    processSaveProductsAssetsJob({
      products: Array.isArray(request.products) ? request.products : [],
      mode: request.mode,
      saveAs: request.saveAs,
    })
      .then(result => sendResponse({ ok: true, ...result }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'save failed' }));
    return true;
  }
  return false;
});
