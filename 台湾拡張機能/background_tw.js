// background.js - ダウンロード制御サービスワーカー

const GAS_SYNC_QUEUE_KEY = 'booksGasSyncQueueV1';
const GAS_SYNC_STATUS_KEY = 'booksGasSyncStatusV1';
const GAS_SYNC_ALARM_NAME = 'booksGasSyncQueueAlarmV1';
const GAS_SYNC_TIMEOUT_MS = 60000;

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
    // GAS returns 302 redirect: script.google.com → script.googleusercontent.com
    // With redirect:'follow' (default), POST becomes GET on 302 and can cause 401.
    // Handle redirects manually to preserve POST method and body.
    let currentUrl = url;
    const MAX_REDIRECTS = 5;

    for (let i = 0; i <= MAX_REDIRECTS; i++) {
      const response = await fetch(currentUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: bodyText,
        credentials: 'omit',
        redirect: 'manual',
        signal: controller.signal,
      });

      if (response.type === 'opaqueredirect' || (response.status >= 300 && response.status < 400)) {
        const location = response.headers.get('location');
        if (location) {
          currentUrl = location;
          continue;
        }
      }

      return {
        status: response.status,
        responseText: await response.text(),
        responseUrl: response.url || currentUrl,
      };
    }

    throw new Error('Too many redirects from GAS');
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
        const appended = Number(result?.appended || result?.results?.length || job.items.length || 0);
        const firstResult = Array.isArray(result?.results) ? result.results[0] : null;
        const locationText = appended === 1 && firstResult?.sheetName && firstResult?.appendedRow
          ? `（${firstResult.sheetName} ${firstResult.appendedRow}行目）`
          : '';

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
            result,
            message: `✅ シートへ ${appended} 件追記しました${locationText}`,
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


function resolveAbsoluteUrl(url, baseUrl = 'https://www.books.com.tw/') {
  try {
    return new URL(String(url || '').trim(), baseUrl).href;
  } catch {
    return String(url || '').trim();
  }
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

async function buildTaiwanImageDownloadUrl(url) {
  const text = String(url || '').trim();
  if (!/^https?:\/\//i.test(text)) return text;
  const blob = await fetchTaiwanImageBlob(text);
  return convertBlobToJpegDataUrl(blob, TAIWAN_IMAGE_MAX_EDGE);
}

chrome.alarms.onAlarm.addListener(alarm => {
  if (alarm?.name === GAS_SYNC_ALARM_NAME) {
    processGasSyncQueue().catch(() => {});
  }
});

function startDownload({ url, filename, saveAs }, sendResponse) {
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

function handleImageDownload(request, sendResponse) {
  const url = String(request.url || '').trim();
  const filename = sanitizeRelativePath(request.filename, '1.jpg');
  if (!url) {
    sendResponse({ ok: false, error: 'url missing' });
    return;
  }

  const saveAs = request.saveAs === true;

  (async () => {
    try {
      const preparedUrl = await buildTaiwanImageDownloadUrl(url);
      startDownload({ url: preparedUrl, filename, saveAs, useFilenameHook: saveAs }, sendResponse);
    } catch (error) {
      sendResponse({ ok: false, error: error?.message || 'image download failed' });
    }
  })();
}

function handleCsvDownload(request, sendResponse) {
  const csvText = String(request.csvText || '');
  const filename = sanitizeRelativePath(request.filename, 'export.csv');
  if (!csvText) {
    sendResponse({ ok: false, error: 'csv text missing' });
    return;
  }

  const csvBytes = new TextEncoder().encode(csvText);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${bytesToBase64(csvBytes)}`;
  startDownload({ url: dataUrl, filename, saveAs: request.saveAs === true }, sendResponse);
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
          buildTaiwanImageDownloadUrl(url).then(preparedUrl => {
            startDownload({ url: preparedUrl, filename, saveAs: false }, () => resolve());
          }).catch(() => resolve());
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

  if (request.action === 'downloadImage') {
    handleImageDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'downloadCsv') {
    handleCsvDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'enqueueDownloads') {
    enqueueDownloadJobs(request, sendResponse);
    return true;
  }

  return false;
});
