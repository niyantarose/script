// background.js - ダウンロード制御サービスワーカー

const GAS_SYNC_QUEUE_KEY = 'booksGasSyncQueueV1';
const GAS_SYNC_STATUS_KEY = 'booksGasSyncStatusV1';
const GAS_SYNC_ALARM_NAME = 'booksGasSyncQueueAlarmV1';
const GAS_SYNC_TIMEOUT_MS = 60000;
const IMAGE_FILENAME_EXT = 'jpg';
const TAIWAN_IMAGE_MAX_EDGE = 1000;
const TAIWAN_DOWNLOAD_MARKER = 'nyantarose_tw___-_';
const pendingTaiwanFilenameByMarker = new Map();
let gasSyncProcessing = false;


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

async function postProductsToGas(url, items) {
  const response = await postGasJson(url, JSON.stringify({
    source: 'books.com.tw-extension',
    appendMode: 'append',
    requestedAt: new Date().toISOString(),
    items,
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

function createGasSyncJob(request) {
  return {
    jobId: `gas_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    url: String(request.url || '').trim(),
    items: Array.isArray(request.items) ? request.items : [],
    requestedAt: new Date().toISOString(),
  };
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

  chrome.alarms.create(GAS_SYNC_ALARM_NAME, { when: Date.now() + 50 });
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
      if (!queue.length) return;

      const job = queue[0];
      const nextQueue = queue.slice(1);

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
        await storageSetLocal({
          [GAS_SYNC_QUEUE_KEY]: nextQueue,
          [GAS_SYNC_STATUS_KEY]: buildGasSyncStatus({
            state: 'error',
            jobId: job.jobId,
            itemCount: job.items.length,
            queueLength: nextQueue.length,
            requestedAt: job.requestedAt,
            finishedAt: new Date().toISOString(),
            error: error?.message || 'GAS write failed',
            message: `❌ ${error?.message || 'GAS write failed'}`,
          }),
        });
      }
    }
  } finally {
    gasSyncProcessing = false;
  }
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



function buildTaiwanMarkerFilename(targetFilename = 'download.bin') {
  const safeTargetFilename = sanitizeRelativePath(targetFilename, 'download.bin');
  const extensionMatch = safeTargetFilename.match(/(\.[a-z0-9]+)$/i);
  const extension = extensionMatch ? extensionMatch[1] : '.bin';
  const token = `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const markerFilename = `${TAIWAN_DOWNLOAD_MARKER}${token}${extension}`;
  pendingTaiwanFilenameByMarker.set(markerFilename, safeTargetFilename);
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
    return direct;
  }

  for (const [markerFilename, desiredFilename] of pendingTaiwanFilenameByMarker.entries()) {
    if (normalized.includes(markerFilename)) {
      pendingTaiwanFilenameByMarker.delete(markerFilename);
      return desiredFilename;
    }
  }

  return '';
}
function isGoodsLikeTaiwanProductCode(productCode = '') {
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

function unwrapTaiwanImageSource(urlText, options = {}) {
  const raw = String(urlText || '').trim();
  if (!raw) return '';

  try {
    const parsed = new URL(raw, 'https://www.books.com.tw/');
    const original = parsed.searchParams.get('i');
    if (original) {
      let originalUrl = decodeURIComponent(original);
      if (originalUrl.startsWith('//')) originalUrl = `https:${originalUrl}`;
      return upgradeTaiwanDirectImageVariant(originalUrl, options);
    }
    return upgradeTaiwanDirectImageVariant(parsed.href, options);
  } catch {
    return upgradeTaiwanDirectImageVariant(raw, options);
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


chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  const desiredFilename = consumePendingTaiwanFilenameByMarker(item?.filename || '');
  if (desiredFilename) {
    suggest({ filename: desiredFilename, conflictAction: 'uniquify' });
    return;
  }
  suggest();
});
chrome.alarms.onAlarm.addListener(alarm => {
  if (alarm?.name === GAS_SYNC_ALARM_NAME) {
    processGasSyncQueue().catch(() => {});
  }
});

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
  const preparedUrl = buildTaiwanImageDownloadUrl(url, {
    keepThumbVariant: isGoodsLikeTaiwanProductCode(productCode),
  });
  const blob = await fetchTaiwanImageBlob(preparedUrl);
  const dataUrl = await convertBlobToJpegDataUrl(blob, TAIWAN_IMAGE_MAX_EDGE);
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
    storageGetLocal([GAS_SYNC_STATUS_KEY])
      .then(result => sendResponse({ ok: true, status: result[GAS_SYNC_STATUS_KEY] || null }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'status read failed' }));
    return true;
  }

  if (request.action === 'clearGasSyncStatus') {
    storageSetLocal({ [GAS_SYNC_STATUS_KEY]: null })
      .then(() => sendResponse({ ok: true }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'status clear failed' }));
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
    })
      .then(result => sendResponse({ ok: true, ...result }))
      .catch(error => sendResponse({ ok: false, error: error.message || 'save failed' }));
    return true;
  }
  return false;
});





