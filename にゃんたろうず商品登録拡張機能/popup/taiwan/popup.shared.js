// 商品リストの管理
let products = [];
let manualSaveImages = false;
let gasWebAppUrl = '';
let gasSyncStatus = null;
let gasSyncPollTimer = null;
let gasUrlDraftTimer = null;

const MANUAL_SAVE_KEY = 'manualSaveImagesV2';
const GAS_URL_KEY = 'booksGasWebAppUrl';
const GAS_URL_DRAFT_KEY = 'booksGasWebAppUrlDraft';
const GAS_SYNC_QUEUE_KEY = 'booksGasSyncQueueV1';
const GAS_SYNC_STATUS_KEY = 'booksGasSyncStatusV1';
const GAS_SYNC_RESET_TOKEN_KEY = 'booksGasSyncResetTokenV1';
const MAX_IMAGE_DOWNLOADS_PER_PRODUCT = 30;
const MAX_FALLBACK_IMAGE_CANDIDATES = 30;
const TARGET_BOOKS_IMAGE_EDGE = 1000;
function isGoodsLikeProductCode(productCode = '') {
  return /^[NM]/i.test(String(productCode || '').trim());
}

function upgradeBooksImageVariant(imageUrl, options = {}) {
  const keepThumbVariant = options.keepThumbVariant === true;
  return String(imageUrl || '')
    .replace(/_t_(\d+)(?=\.(?:jpe?g|png|webp)(?:$|[?#]))/gi, keepThumbVariant ? '_t_$1' : '_b_$1')
    .replace(/\.webp(?=($|[?#]))/i, '.jpeg');
}

const MAIN_IMAGE_ONLY_MODE = false;
const USE_POPUP_FALLBACK_IMAGE_FETCH = false;
const COMIC_BOOK_SHEET_NAME = '台湾まんが';
const OTHER_BOOK_SHEET_NAME = '台湾書籍その他';
const GOODS_SHEET_NAME = '台湾グッズ';
const MAGAZINE_SHEET_NAME = '台湾雑誌';
const MANGA_UPDATES_PROXY_ACTION = 'lookupMangaUpdatesJapaneseTitle';
const MANGA_UPDATES_TITLE_NOT_FOUND = '登録なし';
const MANGA_UPDATES_TITLE_LOOKUP_FAILED = '照会失敗';
const MANGA_UPDATES_TITLE_NOT_LOOKED_UP = '未照会';
const mangaUpdatesTitleCache = new Map();
let mangaUpdatesLookupDisabled = false;
let mangaUpdatesLookupDisabledReason = '';

function storageGet(keys) {
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

function storageSet(items) {
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

function describeSaveDestination(_unusedDirectSaveContext, fallbackLabel = 'Downloads') {
  return fallbackLabel;
}

const BOOK_SHEET_HEADERS = [
  '発番発行',
  '登録状況',
  '商品コード（SKU）',
  'サイト商品コード',
  'タイトル',
  '作者',
  '日本語タイトル',
  'リンク',
  '原題タイトル',
  '原題商品タイトル',
  '売価',
  '原価',
  '粗利益率',
  'ISBN',
  '発売日',
  '言語',
  '単巻数',
  'セット巻数開始番号',
  'セット巻数終了番号',
  'カテゴリ',
  '形態（通常/初回限定/特装）',
  '配送パターン',
  '特典メモ',
  ' メモ',
  'メイン画像',
  '追加画像',
  '予約開始日',
  '予約終了日',
  '入荷予定日',
  '発売日メモ（延期など）',
  '作品ID(W)（自動）',
  'SKU（自動）',
  'ステータス（自動）',
  '残日数（自動）',
  'アラート（自動）',
];

const GOODS_SHEET_HEADERS = [
  '発番発行',
  '登録状況',
  '言語',
  '親コード',
  '商品名（出品用）',
  '売価',
  '原価',
  '粗利益率',
  '配送パターン',
  '商品名（日本語）',
  '日本語タイトル',
  '原題タイトル',
  '特典メモ',
  '商品説明',
  '原題商品タイトル',
  'サイト商品コード',
  'リンク',
  'メイン画像',
  '追加画像',
  '発売日',
  '重複チェックキー',
  '登録日',
  '登録者',
];

const MAGAZINE_SHEET_HEADERS = [
  '登録状況',
  '言語',
  '雑誌名',
  '年',
  '月',
  '表紙情報',
  '特典メモ',
  '親コード',
  'コードステータス',
  '商品名（出品用）',
  '粗利益率',
  '登録日',
  '売価',
  '配送パターン',
  '登録者',
  '商品説明',
  '原価',
  '原題タイトル',
  '原題商品名',
  'アラジン商品コード',
  'アラジンURL',
  '博客來商品コード',
  '博客來URL',
  'メイン画像URL',
  '追加画像URL',
];

function sanitizeFolderName(name) {
  const cleaned = String(name || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
  return cleaned || 'no_title';
}

function buildProductFolderName(product) {
  return sanitizeFolderName(
    extractProductCode(product) || product.ページタイトル || product.商品名 || 'item'
  );
}

function normalizeAdditionalImages(raw) {
  return String(raw || '')
    .replace(/\r?\n+/g, ';')
    .replace(/;{2,}/g, ';')
    .replace(/^;|;$/g, '');
}

function splitImageUrls(raw) {
  return normalizeAdditionalImages(raw)
    .split(';')
    .map(s => s.trim())
    .filter(Boolean);
}

function extractProductCode(product) {
  const pageUrl = String(product?.URL || '').trim();
  return String(product?.商品コード || pageUrl.match(/products\/([^?]+)/)?.[1] || '').trim();
}

function escapeRegExp(str) {
  return String(str || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatToday(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function getBooksImageSourceUrl(imageUrl, productCode = '') {
  let urlText = String(imageUrl || '').trim();
  if (!urlText) return '';
  if (urlText.startsWith('//')) urlText = `https:${urlText}`;

  try {
    const parsed = new URL(urlText, 'https://www.books.com.tw/');
    const original = parsed.searchParams.get('i');
    if (original) {
      let originalUrl = decodeURIComponent(original);
      if (originalUrl.startsWith('//')) originalUrl = `https:${originalUrl}`;
      return upgradeBooksImageVariant(originalUrl, { keepThumbVariant: isGoodsLikeProductCode(productCode) });
    }
    return upgradeBooksImageVariant(parsed.href, { keepThumbVariant: isGoodsLikeProductCode(productCode) });
  } catch {
    return upgradeBooksImageVariant(urlText, { keepThumbVariant: isGoodsLikeProductCode(productCode) });
  }
}
function normalizeBooksImageUrl(rawUrl, productCode = '') {
  return getBooksImageSourceUrl(rawUrl, productCode);
}

function getImageOrderIndexByProductCode(imageUrl, productCode) {
  const code = String(productCode || '');
  if (!code) return Number.MAX_SAFE_INTEGER;

  let basename = '';
  try {
    basename = decodeURIComponent(new URL(getBooksImageSourceUrl(imageUrl, productCode)).pathname.split('/').pop() || '');
  } catch {
    basename = getBooksImageSourceUrl(imageUrl, productCode).split('/').pop() || '';
  }

  const safeCode = escapeRegExp(code);
  const match = basename.match(new RegExp(`^${safeCode}(?:_(?:b|t)_(\\d+))?\\.(?:jpe?g|png|webp)$`, 'i'));
  if (!match) return Number.MAX_SAFE_INTEGER;
  if (!match[1]) return 0;
  const idx = parseInt(match[1], 10);
  return Number.isFinite(idx) ? idx : Number.MAX_SAFE_INTEGER;
}

function isNoiseBooksImageUrl(imageUrl, productCode = '') {
  const text = String(imageUrl || '');
  const baseNoise = /(?:\/G\/ADbanner\/|\/G\/prod\/comingsoon|languageTest|esqsm|\/reco\/|\/banner\/|\/recommend\/|\/ad\/|listE\.php|listN\.php|\/activity\/|\/combo\/|\/event\/|logo|icon)/i;
  if (baseNoise.test(text)) return true;
  return !isGoodsLikeProductCode(productCode) && /_t_\d+\.(?:jpg|jpeg|png|webp)/i.test(text);
}

function isProductBonusBannerImage(imageUrl, productCode = '') {
  const text = String(imageUrl || '');
  const code = String(productCode || '').trim();
  return !!code &&
    new RegExp(escapeRegExp(code), 'i').test(text) &&
    /(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/|\/banner\/)/i.test(text);
}

function buildImageDedupeKey(imageUrl, productCode = '') {
  const normalized = normalizeBooksImageUrl(imageUrl, productCode);
  if (!normalized) return '';

  const idx = getImageOrderIndexByProductCode(normalized, productCode);
  if (idx !== Number.MAX_SAFE_INTEGER) return `slot:${idx}`;

  const sourceUrl = getBooksImageSourceUrl(normalized, productCode);
  try {
    const parsed = new URL(sourceUrl);
    return `${parsed.origin}${parsed.pathname}`.toLowerCase();
  } catch {
    return sourceUrl.replace(/[?#].*$/, '').toLowerCase();
  }
}

function limitImageUrls(urls, max = MAX_IMAGE_DOWNLOADS_PER_PRODUCT, productCode = '') {
  const ordered = [];
  const seen = new Set();

  for (const rawUrl of (urls || [])) {
    const rawText = String(rawUrl || '').trim();
    if (!/^https?:\/\//i.test(rawText)) continue;

    const normalized = normalizeBooksImageUrl(rawText, productCode);
    if (!normalized) continue;
    if (isNoiseBooksImageUrl(normalized, productCode) && !isProductBonusBannerImage(normalized, productCode)) continue;

    const key = buildImageDedupeKey(normalized, productCode);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    ordered.push(normalized);
  }

  const safeMax = Number.isFinite(max) && max > 0 ? Math.floor(max) : MAX_IMAGE_DOWNLOADS_PER_PRODUCT;
  return ordered.slice(0, safeMax);
}

function collectOrderedImageUrls(product) {
  const productCode = extractProductCode(product);
  const mainOnly = limitImageUrls([product.画像URL || ''], 1, productCode);
  if (MAIN_IMAGE_ONLY_MODE) {
    return mainOnly;
  }

  const legacyAdditional = [
    product.画像URL_2枚目 || '',
    product.画像URL_3枚目 || '',
    product.画像URL_4枚目以降 || '',
  ].filter(Boolean).join(';');
  const additional = product.追加画像URL || legacyAdditional;

  const urls = [
    ...mainOnly,
    ...splitImageUrls(additional),
  ].filter(u => /^https?:\/\//i.test(u));

  return limitImageUrls(urls, MAX_IMAGE_DOWNLOADS_PER_PRODUCT, productCode);
}

function isLikelyBooksPageImageUrl(imageUrl, productCode) {
  const text = String(imageUrl || '');
  if (!/^https?:\/\//i.test(text)) return false;
  if (!/(?:\.jpe?g|\.png|\.webp)(?:[?&#]|$)|\/image\/getImage\?|\/fancybox\/getImage\.php\?/i.test(text)) return false;
  if (isNoiseBooksImageUrl(text, productCode) && !isProductBonusBannerImage(text, productCode)) return false;

  if (productCode) {
    return new RegExp(escapeRegExp(productCode), 'i').test(text);
  }
  return /(?:im\d+\.book\.com\.tw\/image\/getImage\?|www\.books\.com\.tw\/img\/)/i.test(text);
}

async function fetchProductPageImageUrls(product) {
  const pageUrl = String(product?.URL || '').trim();
  const productCode = extractProductCode(product);

  if (!USE_POPUP_FALLBACK_IMAGE_FETCH) return [];
  if (!/^https?:\/\/www\.books\.com\.tw\/products\//i.test(pageUrl) || !productCode) {
    return [];
  }

  try {
    const response = await fetch(pageUrl, { credentials: 'omit' });
    if (!response.ok) return [];

    const html = await response.text();
    const candidates = [];
    const pushCandidate = rawUrl => {
      const normalized = normalizeBooksImageUrl(rawUrl, productCode);
      if (isLikelyBooksPageImageUrl(normalized, productCode)) {
        candidates.push(normalized);
      }
    };

    Array.from(html.matchAll(/https?:\/\/[^"'\s<)]+(?:jpe?g|jpeg|png|webp)(?:\?[^"'\s<)]*)?/gi))
      .forEach(match => pushCandidate(match[0]));
    Array.from(html.matchAll(/https?:\/\/[^"'\s<)]*getImage\.php\?[^"'\s<)]*/gi))
      .forEach(match => pushCandidate(match[0]));

    return limitImageUrls(candidates, MAX_FALLBACK_IMAGE_CANDIDATES, productCode);
  } catch {
    return [];
  }
}

function guessImageExt(imageUrl) {
  return 'jpg';
}

function buildImageDownloadPath(titleFolder, seq, ext) {
  return `${titleFolder}/${seq}.jpg`;
}

function buildCsvDownloadPath(filename) {
  return filename;
}

function scheduleGasUrlDraftSave(value, options = {}) {
  const raw = String(value || '').trim();
  gasWebAppUrl = raw;

  if (options.updateUi !== false) {
    updateUI();
  }

  if (gasUrlDraftTimer) {
    clearTimeout(gasUrlDraftTimer);
  }

  gasUrlDraftTimer = setTimeout(() => {
    persistGasUrlValue(raw, options).catch(error => {
      console.error('Failed to persist GAS URL draft:', error);
    });
  }, Number.isFinite(options.delayMs) ? options.delayMs : 120);
}

async function persistGasUrlValue(value, options = {}) {
  const raw = String(value || '').trim();
  gasWebAppUrl = raw;

  const payload = {
    [GAS_URL_DRAFT_KEY]: raw,
  };

  if (options.saveCanonical || isValidGasUrl(raw)) {
    payload[GAS_URL_KEY] = raw;
  }

  await storageSet(payload);
  return raw;
}

function sendRuntimeMessage(message) {
  return new Promise(resolve => {
    chrome.runtime.sendMessage(message, response => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message || 'runtime error' });
        return;
      }
      resolve(response || { ok: false, error: 'empty response' });
    });
  });
}

function getProductSheetName(product) {
  const sheetType = getProductSheetType(product);
  if (sheetType === 'magazine') return MAGAZINE_SHEET_NAME;
  if (sheetType === 'goods') return GOODS_SHEET_NAME;
  return isComicBookProduct(product) ? COMIC_BOOK_SHEET_NAME : OTHER_BOOK_SHEET_NAME;
}

function isGasSyncBusyStatus(status) {
  return !!status?.state && ['queued', 'running'].includes(status.state);
}

function gasSyncStatusType(status) {
  if (!status?.state) return '';
  if (status.state === 'success') return 'success';
  if (status.state === 'error') return 'error';
  return '';
}

function gasSyncStatusMessage(status) {
  if (!status?.state) return '';
  if (status.message) return status.message;
  if (status.state === 'queued') return '⏳ バックグラウンド送信を待機中です';
  if (status.state === 'running') return '📝 バックグラウンドでGASへ書き込み中です';
  if (status.state === 'success') return `✅ シートへ ${Number(status.appended || 0)} 件追記しました`;
  if (status.state === 'error') return `❌ ${status.error || 'GAS write failed'}`;
  return '';
}

async function clearGasSyncStatus() {
  gasSyncStatus = null;
  const response = await sendRuntimeMessage({ action: 'clearGasSyncStatus' });
  if (!response?.ok) {
    throw new Error(response?.error || 'clearGasSyncStatus failed');
  }
  return true;
}

async function resetGasSyncState() {
  gasSyncStatus = null;
  const stored = await storageGet([GAS_SYNC_RESET_TOKEN_KEY]);
  const currentToken = Number(stored[GAS_SYNC_RESET_TOKEN_KEY] || 0);
  const nextToken = Number.isFinite(currentToken) ? currentToken + 1 : 1;
  await storageSet({
    [GAS_SYNC_QUEUE_KEY]: [],
    [GAS_SYNC_STATUS_KEY]: null,
    [GAS_SYNC_RESET_TOKEN_KEY]: nextToken,
  });
  const response = await sendRuntimeMessage({ action: 'resetGasSyncState', resetToken: nextToken });
  if (!response?.ok) {
    console.warn('Failed to notify background about GAS reset:', response?.error || 'resetGasSyncState failed');
  }
  return true;
}

async function refreshGasSyncStatus(options = {}) {
  const response = await sendRuntimeMessage({ action: 'getGasSyncStatus' });
  if (!response?.ok) {
    return options.keepStatusOnError === false ? null : gasSyncStatus;
  }

  gasSyncStatus = response.status || null;
  const message = gasSyncStatusMessage(gasSyncStatus);
  if (message && options.applyMessage !== false) {
    setStatus(message, gasSyncStatusType(gasSyncStatus));
  }

  if (options.updateUi !== false) {
    updateUI();
  }
  return gasSyncStatus;
}

function stopGasSyncStatusPolling() {
  if (!gasSyncPollTimer) return;
  clearInterval(gasSyncPollTimer);
  gasSyncPollTimer = null;
}

function startGasSyncStatusPolling() {
  if (gasSyncPollTimer) return;
  gasSyncPollTimer = setInterval(() => {
    const wasBusy = isGasSyncBusyStatus(gasSyncStatus);
    refreshGasSyncStatus().then(status => {
      if (status?.state === 'success' && typeof scheduleGasSyncSuccessAutoClear === 'function') {
        scheduleGasSyncSuccessAutoClear(status);
      }
      if (!isGasSyncBusyStatus(status)) {
        stopGasSyncStatusPolling();
        // statusがnullになった（スタック状態を解消した）場合、表示を明示的にリセット
        if (!status?.state && wasBusy) {
          setStatus('', '');
        }
      }
    }).catch(() => {});
  }, 1500);
}

function triggerTaiwanAnchorDownload(downloadName, url) {
  const anchor = document.createElement('a');
  anchor.href = String(url || '').trim();
  anchor.download = String(downloadName || '').trim() || 'download.bin';
  anchor.rel = 'noopener';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function waitForQueuedDownloadTick(delayMs = 150) {
  return new Promise(resolve => {
    window.setTimeout(resolve, Math.max(0, Number(delayMs) || 0));
  });
}

async function downloadCsv(csvText, filename) {
  const response = await sendRuntimeMessage({ action: 'prepareTaiwanCsvDownload', csvText, filename });
  if (!response?.ok || !response.url || !response.downloadName) {
    return false;
  }
  triggerTaiwanAnchorDownload(response.downloadName, response.url);
  await waitForQueuedDownloadTick();
  return true;
}

async function downloadImage(imageUrl, filename, saveAs) {
  const response = await sendRuntimeMessage({ action: 'prepareTaiwanImageDownload', url: imageUrl, filename, saveAs: !!saveAs });
  if (!response?.ok || !response.url || !response.downloadName) {
    return false;
  }
  triggerTaiwanAnchorDownload(response.downloadName, response.url);
  await waitForQueuedDownloadTick();
  return true;
}

async function downloadProductImages(items, options = {}) {
  const manualSave = !!options.manualSave;
  let total = 0;
  let success = 0;
  let failed = 0;
  const folders = new Set();

  for (const item of items) {
    const titleFolder = buildProductFolderName(item);
    folders.add(titleFolder);
    let imageUrls = collectOrderedImageUrls(item);
    if (imageUrls.length <= 1) {
      const fallbackUrls = await fetchProductPageImageUrls(item);
      imageUrls = limitImageUrls([...imageUrls, ...fallbackUrls], MAX_IMAGE_DOWNLOADS_PER_PRODUCT, extractProductCode(item));
    }

    let seq = 1;
    for (const imageUrl of imageUrls) {
      const ext = guessImageExt(imageUrl);
      const filePath = buildImageDownloadPath(titleFolder, seq, ext);
      total += 1;
      const ok = await downloadImage(imageUrl, filePath, manualSave);
      if (ok) {
        success += 1;
      } else {
        failed += 1;
      }
      seq += 1;
    }
  }

  return {
    total,
    success,
    failed,
    folders: [...folders],
    apiUnavailable: false,
    folderRuleMisses: 0,
    directSave: false,
    folderName: '',
  };
}

function csvEscape(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}

function buildCsvContent(headers, rows) {
  const bom = String.fromCharCode(0xFEFF);
  return bom + [headers, ...rows]
    .map(row => row.map(csvEscape).join(','))
    .join('\r\n');
}
async function downloadCsvFile(headers, rows, filename, options = {}) {
  const csvContent = buildCsvContent(headers, rows);
  return downloadCsv(csvContent, filename);
}

function trimValue(value) {
  return String(value ?? '').trim();
}

function isGasWebAppExecUrl(url) {
  try {
    const parsed = new URL(trimValue(url));
    return parsed.protocol === 'https:'
      && parsed.hostname === 'script.google.com'
      && /^\/macros\/s\/[^/]+\/exec\/?$/i.test(parsed.pathname);
  } catch {
    return false;
  }
}

function cleanDescription(value) {
  return String(value || '')
    .replace(/\r/g, '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .join('\n');
}

function getAdditionalImagesValue(product) {
  return normalizeAdditionalImages(
    product.追加画像URL || [
      product.画像URL_2枚目 || '',
      product.画像URL_3枚目 || '',
      product.画像URL_4枚目以降 || '',
    ].filter(Boolean).join(';')
  );
}

function extractOriginalTitleText(value) {
  const fallback = trimValue(value);
  if (!fallback) return '';

  const normalized = fallback
    .replace(/^[台臺][湾灣]版\s*/u, '')
    .replace(/\s*【[^】]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^】]*】\s*/gu, ' ')
    .replace(/\s*\([^)]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^)]*\)\s*/gu, ' ')
    .replace(/\s*\[[^\]]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^\]]*\]\s*/gu, ' ')
    .replace(/\s*（[^）]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^）]*）\s*/gu, ' ')
    .replace(/\s*[-/／]\s*(?:首刷|初回|限定|特裝|特装|通常|普通)[^ ]*\s*/gu, ' ')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s+\d+\s*$/u, '')
    .replace(/\s+/g, ' ')
    .trim();

  return normalized || fallback;
}

function detectEditionType(product) {
  const text = [product?.商品名 || '', product?.特典情報 || '', product?.商品説明 || ''].join('\n');
  if (/特裝|特装/.test(text)) return '特装版';
  if (/初回限定|首刷限定|首刷贈品|首刷赠品|買就送|买就送/.test(text)) return '初版限定版';
  if (/限定版/.test(text)) return '限定版';
  return '通常版';
}

function extractBonusMemo(product) {
  return cleanDescription(product?.特典情報 || product?.補足項目 || '');
}

function getCategoryValue(product) {
  const source = [
    product?.カテゴリ,
    product?.商品名,
    product?.ページタイトル,
    product?.商品説明,
  ].map(trimValue).filter(Boolean).join('\n');

  if (!source) return '';
  if (/設定集|設定資料|設定资料|公式設定/.test(source)) return '設定集';
  if (/畫集|画集|畫冊|画冊|イラスト集|美術設定|アートブック/.test(source)) return 'アートブック';
  if (/漫畫|漫画|コミック/.test(source)) return 'まんが';
  if (/小説|小說|小说|輕小說|轻小说|ライトノベル|ノベル|文學|文学/.test(source)) return '小説';
  if (/エッセイ|散文|隨筆|随筆/.test(source)) return 'エッセイ';
  if (/人文社科|社會議題|社会議題|弱勢族群|弱势族群|報導文學|报道文学|紀實|纪实|人物傳記|人物传记|社會觀察|社会观察|議題思辨|议题思辨/.test(source)) return 'エッセイ';
  if (/雜誌|杂志|雑誌/.test(source)) return '雑誌';
  if (/繪本|绘本|絵本/.test(source)) return '絵本';
  if (/シナリオ|scenario|劇本|剧本/i.test(source)) return 'シナリオ集';
  if (/台本/.test(source)) return '台本';
  if (/ステッカー|貼紙|贴纸/.test(source)) return 'ステッカー';
  if (/シール/.test(source)) return 'シール';
  if (/手芸|手藝|手艺/.test(source)) return '手芸';
  if (/切り絵|剪紙|剪纸/.test(source)) return '切り絵';
  if (/Blu-ray|藍光|蓝光|BLAY/i.test(source)) return 'Blu-ray';
  if (/(?:^|[^A-Z])DVD(?:$|[^A-Z])/i.test(source)) return 'DVD';
  if (/(?:^|[^A-Z])CD(?:$|[^A-Z])/i.test(source) || /音樂CD|音乐CD/.test(source)) return 'CD';
  if (/(?:^|[^A-Z])LP(?:$|[^A-Z])/i.test(source)) return 'LP';
  return trimValue(product?.カテゴリ || '');
}

function getLanguageValue(product) {
  const raw = trimValue(product?.言語 || '');
  const source = [raw, trimValue(product?.商品説明 || ''), trimValue(product?.カテゴリ || '')]
    .filter(Boolean)
    .join('\n');

  if (!source) return '台湾';
  if (/繁體|繁体|繁體中文|繁体中文|中文|華文|华文/.test(source)) return '台湾';
  if (/韓文|韩文|韓語|韩语|朝鮮語|朝鲜语/.test(source)) return '韓国';
  if (/日文|日語|日语|日本語/.test(source)) return '日本';
  if (/簡體|简体|中國語|中国语|中文简体|中文簡體/.test(source)) return '中国';
  if (/泰文|タイ語|泰語|泰语/.test(source)) return 'タイ';
  if (/英文|英語|英语|English/i.test(source)) return '英語';
  if (/アメリカ|米国|美国|USA|US\b|American/i.test(source)) return 'アメリカ';
  return '台湾';
}

function getLanguageCodeValue(product) {
  switch (getLanguageValue(product)) {
    case '韓国': return 'KR';
    case '日本': return 'JP';
    case '中国': return 'CN';
    case 'タイ': return 'TH';
    case 'アメリカ': return 'US';
    case '英語': return 'EN';
    case '台湾':
    default:
      return 'TW';
  }
}

function isIgnorableMangaUpdatesError(error) {
  const message = String(error?.message || error || '');
  return /MangaUpdates API blocked extension requests \(403\)/i.test(message)
    || /MangaUpdates API error:\s*403\b/i.test(message)
    || /MangaUpdates proxy error/i.test(message)
    || /MangaUpdates proxy unavailable/i.test(message);
}

function isMangaUpdatesStatusValue(value) {
  const text = trimValue(value);
  return text === MANGA_UPDATES_TITLE_NOT_FOUND
    || text === MANGA_UPDATES_TITLE_LOOKUP_FAILED
    || text === MANGA_UPDATES_TITLE_NOT_LOOKED_UP;
}

function shouldReplaceJapaneseTitleValue(value) {
  const text = trimValue(value);
  return !text || isMangaUpdatesStatusValue(text);
}

function getMangaUpdatesLookupFallbackStatus(queries, hasGasUrl, isLookupTarget) {
  if (mangaUpdatesLookupDisabled && mangaUpdatesLookupDisabledReason === 'failed') {
    return MANGA_UPDATES_TITLE_LOOKUP_FAILED;
  }
  if (!isLookupTarget || !hasGasUrl || !(queries || []).length) {
    return MANGA_UPDATES_TITLE_NOT_LOOKED_UP;
  }
  return MANGA_UPDATES_TITLE_LOOKUP_FAILED;
}

function hasJapaneseTitleSignal(value) {
  return /[ぁ-ゖァ-ヺ々〆ヵヶ]/.test(String(value || ''));
}

function normalizeMangaUpdatesTitleKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000\"'“”‘’・･:：!！?？,，、.．\-ー‐―–—~〜/／\\|()\[\]{}【】「」『』<>]/g, '');
}

function uniqNonEmptyTitles(values) {
  const seen = new Set();
  const list = [];
  for (const value of (values || [])) {
    const text = trimValue(value);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    list.push(text);
  }
  return list;
}

function decodeHtmlEntities(value) {
  return String(value || '')
    .replace(/&#x([0-9a-f]+);/gi, (match, hex) => {
      try {
        return String.fromCodePoint(parseInt(hex, 16));
      } catch {
        return match;
      }
    })
    .replace(/&#(\d+);/g, (match, dec) => {
      try {
        return String.fromCodePoint(parseInt(dec, 10));
      } catch {
        return match;
      }
    })
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function stripMangaUpdatesHtmlToText(html) {
  return decodeHtmlEntities(
    String(html || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(?:p|div|li|tr|section|article|h[1-6])>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
  )
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function extractMangaUpdatesSiteSearchResults(html) {
  const results = [];
  const seenUrls = new Set();
  const pattern = /<a\b[^>]*href="([^"]*(?:\/series\/|series\.html\?id=)[^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
  let match;

  while ((match = pattern.exec(String(html || ''))) !== null) {
    const title = trimValue(stripMangaUpdatesHtmlToText(match[2] || ''));
    if (!title || title.length < 2 || /^Advanced Search$/i.test(title)) continue;

    let url = trimValue(match[1] || '');
    try {
      url = new URL(url, 'https://www.mangaupdates.com/').href;
    } catch {
      // Keep the original href if URL parsing fails.
    }

    if (!/^https:\/\/www\.mangaupdates\.com\//i.test(url) || seenUrls.has(url)) continue;
    seenUrls.add(url);
    results.push({ title, url });
  }

  return results;
}

function extractMangaUpdatesAssociatedTitles(html) {
  const text = stripMangaUpdatesHtmlToText(html);
  const label = 'Associated Names';
  const startIndex = text.indexOf(label);
  if (startIndex < 0) return [];

  let block = text.slice(startIndex + label.length);
  const endLabels = [
    'Groups Scanlating',
    'Latest Release(s)',
    'Recommendations',
    'Reviews',
    'Forum',
    'Status in Country of Origin',
    'Description',
    'Anime Start/End Chapter',
    'Category',
    'Rating',
  ];
  let endIndex = block.length;
  for (const endLabel of endLabels) {
    const index = block.indexOf(endLabel);
    if (index >= 0 && index < endIndex) endIndex = index;
  }
  block = block.slice(0, endIndex);

  return uniqNonEmptyTitles(
    block
      .split(/\n+/)
      .map(title => trimValue(title))
      .filter(title => title && title.length >= 2)
  );
}

function cjkCharOverlap(a, b) {
  const cjkOnly = s => String(s || '').replace(/[^\u3000-\u9fff\uf900-\ufaff]/g, '');
  const charsA = cjkOnly(a);
  const charsB = cjkOnly(b);
  if (!charsA || !charsB) return 0;
  const setA = new Set(charsA);
  const setB = new Set(charsB);
  let common = 0;
  for (const ch of setA) { if (setB.has(ch)) common++; }
  return common / Math.min(setA.size, setB.size);
}

async function fetchMangaUpdatesSiteHtml_(url) {
  const requestUrl = String(url || '').trim();
  if (!requestUrl) {
    return { ok: false, status: 0, body: '', via: 'none', error: 'url missing' };
  }

  const requestOptions = {
    method: 'GET',
    headers: {
      Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    },
    credentials: 'omit',
    cache: 'no-store',
  };

  try {
    const backgroundResp = await sendRuntimeMessage({
      action: 'mangaUpdatesApiFetch',
      url: requestUrl,
      options: requestOptions,
    });
    if (backgroundResp?.ok || Number(backgroundResp?.status || 0) > 0 || String(backgroundResp?.body || '')) {
      return {
        ok: !!backgroundResp?.ok,
        status: Number(backgroundResp?.status || 0),
        body: String(backgroundResp?.body || ''),
        via: 'background',
        error: String(backgroundResp?.error || ''),
      };
    }
  } catch (error) {
    console.warn('[MU直接] background fetch failed:', error?.message || error);
  }

  try {
    const directResp = await fetch(requestUrl, requestOptions);
    const body = await directResp.text().catch(() => '');
    return {
      ok: directResp.ok,
      status: directResp.status,
      body,
      via: 'direct',
      error: '',
    };
  } catch (error) {
    return {
      ok: false,
      status: 0,
      body: '',
      via: 'direct',
      error: String(error?.message || error || ''),
    };
  }
}

async function fetchMangaUpdatesJapaneseTitleDirect(queries) {
  const MU_SITE_SEARCH = 'https://www.mangaupdates.com/site/search/result?search=';
  const normalizedQueries = uniqNonEmptyTitles(queries || []).slice(0, 8);
  if (!normalizedQueries.length) return { ok: true, japaneseTitle: '' };
  console.log('[MU直接] 検索開始:', normalizedQueries);

  for (const query of normalizedQueries) {
    try {
      const searchResp = await fetchMangaUpdatesSiteHtml_(MU_SITE_SEARCH + encodeURIComponent(query));
      console.log('[MU直接] site search status:', searchResp?.status, 'via:', searchResp?.via, 'query:', query);
      if (!searchResp?.ok) continue;
      const searchHtml = searchResp.body || '';
      const results = extractMangaUpdatesSiteSearchResults(searchHtml).slice(0, 5);

      for (let ri = 0; ri < results.length; ri++) {
        const record = results[ri];
        let candidateTitles = uniqNonEmptyTitles([record?.title || '']);
        if (!record?.url) continue;

        const detailResp = await fetchMangaUpdatesSiteHtml_(record.url);
        if (detailResp?.ok) {
          const detailHtml = detailResp.body || '';
          candidateTitles = uniqNonEmptyTitles([
            record.title,
            ...extractMangaUpdatesAssociatedTitles(detailHtml),
          ]);
        }

        const queryKey = normalizeMangaUpdatesTitleKey(query);
        if (!queryKey || queryKey.length < 2) continue;
        const matched = candidateTitles.some(c => {
          const ck = normalizeMangaUpdatesTitleKey(c);
          if (!ck || ck.length < 2) return false;
          if (ck === queryKey) return true;
          if (ck.length >= 4 && queryKey.length >= 4 && (ck.includes(queryKey) || queryKey.includes(ck))) return true;
          if (ri < 3 && cjkCharOverlap(query, c) >= 0.5) return true;
          return false;
        });
        console.log('[MU直接] series:', record.title, 'matched:', matched);
        if (!matched) continue;

        for (const c of candidateTitles) {
          if (hasJapaneseTitleSignal(c)) {
            console.log('[MU直接] 日本語タイトル発見:', c);
            return { ok: true, japaneseTitle: trimValue(c) };
          }
        }
      }
    } catch (e) {
      console.warn('[MU直接] エラー:', e.message);
    }
  }

  console.log('[MU直接] 見つからず');
  return { ok: true, japaneseTitle: '' };
}

function getMangaUpdatesGenreText(product) {
  return [
    product?.ジャンル,
    product?.カテゴリ,
    product?.categoryName,
    product?.mallType,
    typeof getProductSheetType === 'function' ? getProductSheetType(product) : '',
  ].map(trimValue).filter(Boolean).join(' ');
}

function buildMangaUpdatesLookupQueries(product) {
  const rawTitle = trimValue(product?.商品名 || product?.title || product?.name || '');
  const goodsWorkTitle = typeof extractGoodsWorkTitleText === 'function'
    ? trimValue(product?.作品名原題 || extractGoodsWorkTitleText(rawTitle))
    : trimValue(product?.作品名原題 || '');

  return uniqNonEmptyTitles([
    trimValue(product?.原題タイトル || ''),
    goodsWorkTitle,
    trimValue(product?.作品名原題 || ''),
    trimValue(product?.原題商品名 || ''),
    trimValue(product?.['商品名（原題）'] || ''),
    extractOriginalTitleText(rawTitle),
  ]).filter(title => normalizeMangaUpdatesTitleKey(title).length >= 2);
}

function shouldUseMangaUpdatesLookup(product) {
  if (mangaUpdatesLookupDisabled) return false;
  if (!product) return false;

  const hasSourceTitle = buildMangaUpdatesLookupQueries(product).length > 0;
  if (!hasSourceTitle) return false;

  return true;
}

async function lookupJapaneseTitleViaMangaUpdates(query) {
  if (mangaUpdatesLookupDisabled) return '';

  const cacheKey = normalizeMangaUpdatesTitleKey(query);
  if (!cacheKey) return '';
  if (mangaUpdatesTitleCache.has(cacheKey)) return mangaUpdatesTitleCache.get(cacheKey) || '';

  const directResult = await fetchMangaUpdatesJapaneseTitleDirect([query]);
  const japaneseTitle = trimValue(directResult?.japaneseTitle || '');
  mangaUpdatesTitleCache.set(cacheKey, japaneseTitle);
  return japaneseTitle;
}

async function enrichProductWithMangaUpdatesJapaneseTitle(product) {
  if (!product) return product;

  const sheetType = typeof getProductSheetType === 'function' ? getProductSheetType(product) : '';
  const currentJapaneseTitle = trimValue(product?.日本語タイトル || '');
  const needsJapaneseTitle = shouldReplaceJapaneseTitleValue(currentJapaneseTitle);
  const needsWorkJapaneseTitle = sheetType === 'goods'
    && !trimValue(product?.作品名日本語 || product?.['作品名（日本語）'] || '');
  if (!needsJapaneseTitle && !needsWorkJapaneseTitle) return product;

  const queries = buildMangaUpdatesLookupQueries(product);
  const isLookupTarget = queries.length > 0;

  if (!isLookupTarget || !queries.length || mangaUpdatesLookupDisabled) {
    if (needsJapaneseTitle && shouldReplaceJapaneseTitleValue(product?.日本語タイトル || '')) {
      product.日本語タイトル = getMangaUpdatesLookupFallbackStatus(queries, true, isLookupTarget);
    }
    return product;
  }

  let resolvedJapaneseTitle = '';
  for (const query of queries) {
    let japaneseTitle = '';
    try {
      japaneseTitle = await lookupJapaneseTitleViaMangaUpdates(query);
    } catch (error) {
      if (isIgnorableMangaUpdatesError(error)) {
        mangaUpdatesLookupDisabled = true;
        mangaUpdatesLookupDisabledReason = 'failed';
        if (needsJapaneseTitle && shouldReplaceJapaneseTitleValue(product?.日本語タイトル || '')) {
          product.日本語タイトル = MANGA_UPDATES_TITLE_LOOKUP_FAILED;
        }
        return product;
      }
      throw error;
    }
    if (!japaneseTitle) continue;

    resolvedJapaneseTitle = japaneseTitle;
    if (needsJapaneseTitle && shouldReplaceJapaneseTitleValue(product?.日本語タイトル || '')) {
      product.日本語タイトル = japaneseTitle;
    }
    if (needsWorkJapaneseTitle && !trimValue(product?.作品名日本語 || product?.['作品名（日本語）'] || '')) {
      product.作品名日本語 = japaneseTitle;
    }
    break;
  }

  if (!resolvedJapaneseTitle && needsJapaneseTitle && shouldReplaceJapaneseTitleValue(product?.日本語タイトル || '')) {
    product.日本語タイトル = MANGA_UPDATES_TITLE_NOT_FOUND;
  }

  return product;
}

function sendTabMessage(tabId, message) {
  return new Promise(resolve => {
    chrome.tabs.sendMessage(tabId, message, response => {
      resolve({
        response,
        lastError: chrome.runtime.lastError?.message || ''
      });
    });
  });
}
async function requestActiveTabProductInfo() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url?.includes('books.com.tw/products/')) {
    throw new Error('books.com.tw の商品ページを開いてください');
  }

  let { response, lastError } = await sendTabMessage(tab.id, { action: 'getProductInfo' });
  if (lastError) {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['content/taiwan/content.js']
    });
    await new Promise(resolve => setTimeout(resolve, 500));
    ({ response, lastError } = await sendTabMessage(tab.id, { action: 'getProductInfo' }));
  }

  if (lastError || !response || !response.success) {
    throw new Error(response?.error || lastError || '不明なエラー');
  }

  return response.data;
}




























