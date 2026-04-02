// background.js - ダウンロード制御サービスワーカー

const pendingFilenameById = new Map();
const pendingFilenameByUrl = new Map();
const LAST_STATUS_KEY = 'kr_goods_last_status';

function sanitizePathSegment(segment) {
  return String(segment || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
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
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

function setLastStatus(message, type = '') {
  chrome.storage.local.set({
    [LAST_STATUS_KEY]: {
      message: String(message || ''),
      type: String(type || ''),
      updatedAt: Date.now(),
    }
  });
}

function isImageUrl(url) {
  const text = String(url || '').trim();
  if (!/^https?:\/\//i.test(text)) return false;
  if (/image_zoom\d?\.html|\/product\/detail\.html/i.test(text)) return false;
  if (/\.html?(\?|$)/i.test(text)) return false;
  return /\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?|$)/i.test(text);
}

function getImageExt(url) {
  const ext = (String(url || '').split('.').pop() || '').split('?')[0].toLowerCase();
  if (ext === 'jpeg') return 'jpg';
  return ['jpg', 'png', 'webp', 'gif', 'bmp', 'avif'].includes(ext) ? ext : 'jpg';
}

function buildShopProductCode(product) {
  const code = String(product?.ショップコード || product?.ショップ || 'SHOP').trim();
  const id = String(product?.商品ID || '').trim();
  return id ? `${code}-${id}` : code;
}

function buildExportBaseName(products, datePart, timePart) {
  const first = Array.isArray(products) && products.length > 0 ? products[0] : {};
  const shopName = sanitizePathSegment(first.ショップ || first.ショップコード || 'SHOP');
  const shopProductCode = sanitizePathSegment(buildShopProductCode(first));
  const fallback = `SHOP_${datePart || ''}_${timePart || ''}`;
  if (shopName && shopProductCode) return `${shopName}_${shopProductCode}`;
  if (shopName) return shopName;
  if (shopProductCode) return shopProductCode;
  return sanitizePathSegment(fallback) || 'SHOP_EXPORT';
}

function buildCsvContent(products) {
  const headers = [
    'ショップコード', '商品ID', '重複チェックキー',
    '商品名', '価格', 'メイン画像URL', '追加画像URL',
    '商品説明', '購入URL'
  ];
  const rows = (products || []).map(p => [
    p.ショップコード || '', p.商品ID || '', p.重複チェックキー || '',
    p.商品名 || '', p.価格 || '', p.メイン画像URL || '',
    p.追加画像URL || '', p.商品説明 || '', p.購入URL || '',
  ]);

  return '\uFEFF' + [headers, ...rows]
    .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
    .join('\r\n');
}

function toAbsoluteUrl(url, baseUrl) {
  const text = String(url || '').trim();
  if (!text) return '';
  try {
    return new URL(text, baseUrl).href;
  } catch (_e) {
    return '';
  }
}

function getDedupeKey(url) {
  return String(url || '').replace(/[?#].*$/, '');
}

function isNnEditorImageUrl(url) {
  return /\/web\/upload\/NNEditor\//i.test(String(url || ''));
}

function extractAdditionalImageUrlsFromHtml(html, pageUrl, mainImageUrl) {
  const tags = String(html || '').match(/<img\b[^>]*>/gi) || [];
  const urls = [];
  const seen = new Set();
  const mainKey = getDedupeKey(mainImageUrl);
  const isAllowedDetailPath = url => /\/web\/upload\/NNEditor\//i.test(String(url || ''));

  const readUrlsFromTags = (tagList, { requireNnEditor = false } = {}) => {
    for (const tag of tagList) {
      if (!/(ec-data-src|data-ec-data-src|data-src|data-lazy-src|data-original|src)\s*=/i.test(tag)) continue;

      const styleMatch = tag.match(/style\s*=\s*["']([^"']*)["']/i);
      const styleText = (styleMatch?.[1] || '').toLowerCase();
      const minHeightMatch = styleText.match(/min-height\s*:\s*(\d+(?:\.\d+)?)px/);
      if (minHeightMatch) {
        const minHeight = parseFloat(minHeightMatch[1]);
        if (!Number.isNaN(minHeight) && minHeight > 0 && minHeight < 180) continue;
      }

      const valueMatch = tag.match(/(?:ec-data-src|data-ec-data-src|data-src|data-lazy-src|data-original|src)\s*=\s*["']([^"']+)["']/i);
      if (!valueMatch?.[1]) continue;
      const abs = toAbsoluteUrl(valueMatch[1], pageUrl);
      if (!isImageUrl(abs)) continue;
      if (/\/(category|icon|common|layout|banner)\//i.test(abs)) continue;
      if (requireNnEditor && !isAllowedDetailPath(abs)) continue;

      const key = getDedupeKey(abs);
      if (!key || key === mainKey || seen.has(key)) continue;
      seen.add(key);
      urls.push(abs);
    }
  };

  const ecTags = tags.filter(tag => /(ec-data-src|data-ec-data-src)\s*=/i.test(tag));
  if (ecTags.length > 0) {
    // ec-data-src がある場合はまずそれを優先（NNEditor限定にしない）
    readUrlsFromTags(ecTags, { requireNnEditor: false });
    if (urls.length > 0) return urls;
  }

  // まずNNEditorを優先し、0件なら一般画像へフォールバック
  readUrlsFromTags(tags, { requireNnEditor: true });
  if (urls.length === 0) {
    readUrlsFromTags(tags, { requireNnEditor: false });
  }
  return urls;
}

async function resolveAdditionalUrlsForProduct(product) {
  const mainKey = getDedupeKey(product?.メイン画像URL || '');
  const fromDataAll = String(product?.追加画像URL || '')
    .split(';')
    .map(u => String(u || '').trim())
    .filter(isImageUrl)
    .filter(url => !/\/(category|icon|common|layout|banner)\//i.test(url))
    .filter(url => getDedupeKey(url) !== mainKey);

  const fromDataNnEditor = fromDataAll.filter(isNnEditorImageUrl);
  if (fromDataNnEditor.length > 0) {
    return fromDataNnEditor;
  }
  if (fromDataAll.length > 0) {
    return fromDataAll;
  }

  const pageUrl = String(product?.購入URL || '').trim();
  let fromHtml = [];

  if (/^https?:\/\//i.test(pageUrl)) {
    try {
      const resp = await fetch(pageUrl, {
        method: 'GET',
        credentials: 'include',
        referrer: pageUrl,
        referrerPolicy: 'strict-origin-when-cross-origin',
        cache: 'no-store',
      });
      if (resp.ok) {
        const html = await resp.text();
        fromHtml = extractAdditionalImageUrlsFromHtml(html, pageUrl, product?.メイン画像URL || '');
      }
    } catch (_e) {
      // no-op
    }
  }

  const fromHtmlNnEditor = fromHtml.filter(isNnEditorImageUrl);
  if (fromHtmlNnEditor.length > 0) {
    return fromHtmlNnEditor;
  }
  return fromHtml;
}

function formatTopFailReasons(reasonCounts) {
  if (!reasonCounts || reasonCounts.size === 0) return '';
  const items = Array.from(reasonCounts.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 2)
    .map(([reason, count]) => `${reason}:${count}`);
  return items.join(', ');
}

function normalizeFailReason(errMsg) {
  const msg = String(errMsg || '').toLowerCase();
  if (!msg) return 'unknown';
  if (msg.includes('html response')) return 'html';
  if (msg.includes('403')) return 'http403';
  if (msg.includes('404')) return 'http404';
  if (msg.includes('network')) return 'network';
  if (msg.includes('invalid image')) return 'invalid-url';
  return msg.slice(0, 24);
}

async function runExportJob(request) {
  const products = Array.isArray(request?.products) ? request.products : [];
  const datePart = String(request?.datePart || '');
  const timePart = String(request?.timePart || '');

  if (products.length === 0) {
    setLastStatus('❌ 出力失敗: 商品がありません', 'error');
    return;
  }

  setLastStatus('⏳ バックグラウンドで出力中です...', '');

  const exportBaseName = buildExportBaseName(products, datePart, timePart);
  const csvFilename = `${exportBaseName}.csv`;
  const csvContent = buildCsvContent(products);
  const csvRes = await downloadCsvCore({ csvText: csvContent, filename: csvFilename, saveAs: false });
  if (!csvRes?.ok) {
    setLastStatus(`❌ CSV保存失敗: ${csvRes?.error || 'unknown error'}`, 'error');
    return;
  }

  let imageSuccess = 0;
  let imageFailed = 0;
  let imageAttempted = 0;
  const failReasonCounts = new Map();

  for (const p of products) {
    const shopName = sanitizePathSegment(p?.ショップ || p?.ショップコード || 'SHOP');
    const shopProductCode = sanitizePathSegment(buildShopProductCode(p));
    const productFolder = (shopName && shopProductCode)
      ? `${shopName}_${shopProductCode}`
      : (shopName || shopProductCode || sanitizePathSegment(p?.商品名 || '') || 'item');
    const additionalUrls = await resolveAdditionalUrlsForProduct(p);
    const allImages = [
      p?.メイン画像URL,
      ...additionalUrls,
    ]
      .map(url => String(url || '').trim())
      .filter(isImageUrl);

    let imageSeq = 1;
    for (const url of allImages) {
      imageAttempted += 1;
      const ext = getImageExt(url);
      const filename = `${productFolder}/${imageSeq}.${ext}`;
      imageSeq += 1;
      const res = await downloadImageCore({
        url,
        filename,
        saveAs: false,
        referer: String(p?.購入URL || '').trim(),
      });
      if (res?.ok) {
        imageSuccess += 1;
      } else {
        imageFailed += 1;
        const reason = normalizeFailReason(res?.error);
        failReasonCounts.set(reason, (failReasonCounts.get(reason) || 0) + 1);
      }
    }
  }

  const reasonSummary = formatTopFailReasons(failReasonCounts);
  setLastStatus(
    `✅ CSV: ${csvFilename} / 画像: ${imageSuccess}/${imageAttempted}枚${imageFailed > 0 ? ` (失敗:${imageFailed}${reasonSummary ? ` ${reasonSummary}` : ''})` : ''}`,
    imageFailed > 0 ? 'error' : 'success'
  );
}

function buildReferer(sourceUrl, requestReferer) {
  const custom = String(requestReferer || '').trim();
  if (/^https?:\/\//i.test(custom)) {
    return custom;
  }
  try {
    return new URL(sourceUrl).origin + '/';
  } catch (_e) {
    return '';
  }
}

async function fetchAsDataUrl(sourceUrl, fallbackMime, requestReferer) {
  const referer = buildReferer(sourceUrl, requestReferer);
  const headers = {
    'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
  };

  // Refererを付けてfetch（cafe24などReferer必須サイト対応）
  const response = await fetch(sourceUrl, {
    headers,
    credentials: 'include',
    referrer: referer || undefined,
    referrerPolicy: 'strict-origin-when-cross-origin',
    cache: 'no-store',
  });
  if (!response.ok) {
    const err = new Error(`fetch ${response.status}`);
    err.code = `HTTP_${response.status}`;
    throw err;
  }

  const contentType = response.headers.get('content-type') || fallbackMime || 'application/octet-stream';

  // HTMLが返ってきた場合（リダイレクト・エラーページ）はスキップ
  if (contentType.includes('text/html')) {
    const err = new Error(`html response (not an image): ${contentType}`);
    err.code = 'HTML_RESPONSE';
    throw err;
  }

  const bytes = new Uint8Array(await response.arrayBuffer());
  const base64 = bytesToBase64(bytes);
  return `data:${contentType};base64,${base64}`;
}

function consumePendingFilename(item) {
  if (!item) return '';

  const byId = pendingFilenameById.get(item.id);
  if (byId) {
    pendingFilenameById.delete(item.id);
    const candidates = [item.finalUrl, item.url].filter(Boolean);
    candidates.forEach(url => pendingFilenameByUrl.delete(url));
    return byId;
  }

  const candidates = [item.finalUrl, item.url].filter(Boolean);
  for (const url of candidates) {
    if (pendingFilenameByUrl.has(url)) {
      const filename = pendingFilenameByUrl.get(url);
      pendingFilenameByUrl.delete(url);
      return filename;
    }
  }

  return '';
}

chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  // 自分の拡張機能が開始したダウンロードだけ命名する（他拡張/通常DLとは競合しない）
  if (item.byExtensionId !== chrome.runtime.id) {
    suggest();
    return;
  }

  // 非同期で遅延させて最終suggestを取り、他拡張の誤上書きに負けにくくする
  setTimeout(() => {
    const desiredFilename = consumePendingFilename(item);
    if (desiredFilename) {
      suggest({ filename: desiredFilename, conflictAction: 'uniquify' });
      return;
    }
    suggest();
  }, 150);

  return true;
});

chrome.downloads.onChanged.addListener(delta => {
  if (!delta || typeof delta.id !== 'number' || !delta.state) return;
  const state = delta.state.current;
  if (state === 'complete' || state === 'interrupted') {
    pendingFilenameById.delete(delta.id);
  }
});

function startDownload({ url, filename, saveAs }, sendResponse) {
  if (!url) {
    sendResponse({ ok: false, error: 'download url missing' });
    return;
  }

  const safeFilename = sanitizeRelativePath(filename, 'download.bin');
  const useSaveAs = saveAs !== false;

  pendingFilenameByUrl.set(url, safeFilename);

  chrome.downloads.download(
    {
      url,
      filename: safeFilename,
      saveAs: useSaveAs,
      conflictAction: 'uniquify',
    },
    downloadId => {
      if (chrome.runtime.lastError || !downloadId) {
        pendingFilenameByUrl.delete(url);
        sendResponse({ ok: false, error: chrome.runtime.lastError?.message || 'download failed' });
        return;
      }

      pendingFilenameById.set(downloadId, safeFilename);
      sendResponse({ ok: true, downloadId, filename: safeFilename });
    }
  );
}

function startDownloadAsync(params) {
  return new Promise(resolve => {
    startDownload(params, res => resolve(res || { ok: false, error: 'download failed' }));
  });
}

async function downloadCsvCore({ csvText, filename, saveAs = false }) {
  const text = String(csvText || '');
  if (!text) return { ok: false, error: 'csv text missing' };

  const safeFilename = sanitizeRelativePath(filename, 'export.csv');
  const csvBytes = new TextEncoder().encode(text);
  const csvBase64 = bytesToBase64(csvBytes);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${csvBase64}`;
  return await startDownloadAsync({ url: dataUrl, filename: safeFilename, saveAs });
}

async function handleImageDownload(request, sendResponse) {
  const sourceUrl = String(request.url || '').trim();
  const filename = sanitizeRelativePath(request.filename, '1.jpg');
  const saveAs = request.saveAs !== false;
  const referer = String(request.referer || '').trim();

  if (!sourceUrl) {
    sendResponse({ ok: false, error: 'image url missing' });
    return;
  }

  if (/image_zoom\d?\.html|\/product\/detail\.html/i.test(sourceUrl) || /\.html?(\?|$)/i.test(sourceUrl)) {
    sendResponse({ ok: false, error: 'non-image url blocked' });
    return;
  }

  const isLikelyImageUrl = /\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?|$)/i.test(sourceUrl);
  if (!isLikelyImageUrl) {
    sendResponse({ ok: false, error: 'invalid image extension' });
    return;
  }

  try {
    const dataUrl = await fetchAsDataUrl(sourceUrl, 'image/jpeg', referer);
    startDownload({ url: dataUrl, filename, saveAs }, sendResponse);
  } catch (error) {
    // fetch不可(CORS/ネットワーク)のみ直接DLへフォールバック
    if (error?.name === 'TypeError') {
      startDownload({ url: sourceUrl, filename, saveAs }, sendResponse);
      return;
    }
    sendResponse({ ok: false, error: error?.message || 'image fetch failed' });
  }
}

function handleCsvDownload(request, sendResponse) {
  const csvText = String(request.csvText || '');
  const filename = sanitizeRelativePath(request.filename, 'export.csv');
  const saveAs = request.saveAs !== false;

  if (!csvText) {
    sendResponse({ ok: false, error: 'csv text missing' });
    return;
  }

  const csvBytes = new TextEncoder().encode(csvText);
  const csvBase64 = bytesToBase64(csvBytes);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${csvBase64}`;
  startDownload({ url: dataUrl, filename, saveAs }, sendResponse);
}

async function downloadImageCore({ url, filename, saveAs = false, referer = '' }) {
  const sourceUrl = String(url || '').trim();
  const safeFilename = sanitizeRelativePath(filename, '1.jpg');

  if (!isImageUrl(sourceUrl)) {
    return { ok: false, error: 'invalid image url' };
  }

  try {
    const dataUrl = await fetchAsDataUrl(sourceUrl, 'image/jpeg', referer);
    return await startDownloadAsync({ url: dataUrl, filename: safeFilename, saveAs });
  } catch (error) {
    // fetch不可(CORS/ネットワーク)のみ直接DLへフォールバック
    if (error?.name === 'TypeError') {
      return await startDownloadAsync({ url: sourceUrl, filename: safeFilename, saveAs });
    }
    return { ok: false, error: error?.message || 'image fetch failed' };
  }
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'startExport') {
    runExportJob(request)
      .catch(err => setLastStatus(`❌ 出力失敗: ${err?.message || String(err)}`, 'error'));
    sendResponse({ ok: true, started: true });
    return false;
  }

  if (request.action === 'downloadImage') {
    handleImageDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'downloadCsv') {
    handleCsvDownload(request, sendResponse);
    return true;
  }

  sendResponse({ ok: false, error: 'unknown action' });
  return false;
});

