// background.js - アラジン拡張
// ダウンロードとシート書込をバックグラウンドで継続実行する

const STORAGE_KEY = 'aladin_products';
const BACKGROUND_STATUS_STORAGE = 'aladin_background_status';
const JOB_QUEUE_STORAGE = 'aladin_job_queue';
const JOB_ALARM_NAME = 'aladin-job-queue';

let isProcessingQueue = false;


function storageGet(keys) {
  return new Promise(resolve => {
    chrome.storage.local.get(keys, result => {
      resolve(result || {});
    });
  });
}

function storageSet(items) {
  return new Promise(resolve => {
    chrome.storage.local.set(items, () => resolve());
  });
}

function storageRemove(keys) {
  return new Promise(resolve => {
    chrome.storage.local.remove(keys, () => resolve());
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
  startDownload({ url, filename, saveAs }, sendResponse);
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

function downloadImageTask(url, filename, saveAs = false) {
  return new Promise(resolve => {
    handleImageDownload({ url, filename, saveAs }, response => {
      resolve(response?.ok ? { ok: true } : { ok: false, error: response?.error || 'download failed' });
    });
  });
}

function downloadCsvTask(csvText, filename, saveAs = false) {
  return new Promise(resolve => {
    handleCsvDownload({ csvText, filename, saveAs }, response => {
      resolve(Boolean(response?.ok));
    });
  });
}

function detectGenre(item) {
  const categoryName = String(item?.categoryName || '').toLowerCase();
  const mallType = String(item?.mallType || '').toLowerCase();
  const title = String(item?.title || '').toLowerCase();
  const subtitle = String(item?.subInfo?.subTitle || '').toLowerCase();
  const combined = `${categoryName} ${mallType} ${title} ${subtitle}`;

  const hasMagazineKeyword = /magazine|잡지|매거진/i.test(combined);
  const hasMagazineIssuePattern = /\b20\d{2}[./-]\d{1,2}\b/.test(title) || /\b\d{4}年\s*\d{1,2}月\b/.test(title) || /\b\d{4}년\s*\d{1,2}월\b/.test(title);
  const hasMagazineBrand = /(elle|vogue|gq|esquire|allure|bazaar|harper'?s bazaar|marie claire|cosmopolitan|dazed|arena|w korea|ceci|1st look|cine21|maxim|singles|nylon|the star|star1|men'?s health)/i.test(title);

  if (hasMagazineKeyword || (hasMagazineIssuePattern && hasMagazineBrand)) return '雑誌';
  if (/music|음반|cd|lp/i.test(combined)) return '音楽映像';
  if (/dvd|bluray|blu-ray|블루레이|video/i.test(combined)) return '音楽映像';
  if (/comic|만화/i.test(combined)) return 'マンガ';
  if (/gift|goods|굿즈/i.test(combined)) return 'グッズ';
  if (/book|도서|novel|소설|essay|에세이/i.test(combined)) return '書籍';
  return '書籍';
}

const MAGAZINE_NAME_PATTERNS = [
  { name: 'DAZED & CONFUSED', pattern: /dazed\s*&\s*confused/i },
  { name: "HARPER'S BAZAAR", pattern: /harper'?s\s*bazaar/i },
  { name: 'BAZAAR', pattern: /\bbazaar\b/i },
  { name: 'MARIE CLAIRE', pattern: /marie\s*claire/i },
  { name: 'COSMOPOLITAN', pattern: /cosmopolitan/i },
  { name: 'ESQUIRE', pattern: /esquire/i },
  { name: 'ALLURE', pattern: /allure/i },
  { name: 'VOGUE', pattern: /vogue/i },
  { name: 'ELLE', pattern: /elle/i },
  { name: 'GQ', pattern: /\bgq\b/i },
  { name: 'ARENA', pattern: /arena/i },
  { name: 'W KOREA', pattern: /\bw\s*korea\b/i },
  { name: 'CECI', pattern: /\bceci\b/i },
  { name: '1ST LOOK', pattern: /1st\s*look/i },
  { name: 'CINE21', pattern: /cine\s*21/i },
  { name: 'MAXIM', pattern: /maxim/i },
  { name: 'SINGLES', pattern: /singles/i },
  { name: 'NYLON', pattern: /nylon/i },
  { name: 'THE STAR', pattern: /the\s*star/i },
  { name: 'STAR1', pattern: /star\s*1|star1/i },
  { name: "MEN'S HEALTH", pattern: /men'?s\s*health/i }
];

function extractMagazineName(itemLike) {
  const title = String(itemLike?.title || itemLike?.name || itemLike || '').trim();
  const categoryName = String(itemLike?.categoryName || '').trim();
  const combined = `${title} ${categoryName}`;

  for (const entry of MAGAZINE_NAME_PATTERNS) {
    if (entry.pattern.test(combined)) {
      return entry.name;
    }
  }

  const cleaned = title
    .replace(/\b20\d{2}[./-]\d{1,2}\b.*$/i, '')
    .replace(/\b\d{4}년\s*\d{1,2}월.*$/i, '')
    .replace(/\b\d{4}年\s*\d{1,2}月.*$/i, '')
    .replace(/\s+(KOREA|코리아)\b/ig, '')
    .replace(/[<＜【\[(].*$/, '')
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned || title;
}

function buildCsvContent(product) {
  const toCsv = (headers, row) =>
    '\uFEFF' + [headers, row]
      .map(record => record.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\r\n');

  const genre = product.ジャンル;
  const additional = product.追加画像URL || '';
  const description = product.description || '';

  if (genre === '音楽映像') {
    return toCsv(
      ['商品名（原題）', 'アーティスト', 'レーベル', '発売日', '原価', 'ジャンル分類', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, product.author, product.publisher, product.pubDate,
       product.priceSales, product.categoryName, product.cover, additional,
       description, product.pageUrl]
    );
  }
  if (genre === '雑誌') {
    const magazineName = product.magazineName || extractMagazineName(product);
    return toCsv(
      ['原題タイトル', '原題商品名', '原価', '商品説明', 'アラジン商品コード', 'アラジンURL', 'メイン画像URL', '追加画像URL'],
      [magazineName || product.title, product.title, product.priceSales, description,
       product.itemId, product.pageUrl, product.cover, additional]
    );
  }
  if (genre === 'マンガ') {
    return toCsv(
      ['商品名（原題）', '著者', '出版社', '発売日', '原価', 'ジャンル分類', 'ISBN', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, product.author, product.publisher, product.pubDate,
       product.priceSales, product.categoryName, product.isbn13,
       product.cover, additional, description, product.pageUrl]
    );
  }
  if (genre === 'グッズ') {
    return toCsv(
      ['商品名（原題）', '発売日', '原価', 'ジャンル分類', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, product.pubDate, product.priceSales, product.categoryName,
       product.cover, additional, description, product.pageUrl]
    );
  }

  return toCsv(
    ['商品名（原題）', '著者', '出版社', '発売日', '原価', 'ジャンル分類', 'ISBN', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
    [product.title, product.author, product.publisher, product.pubDate,
     product.priceSales, product.categoryName, product.isbn13,
     product.cover, additional, description, product.pageUrl]
  );
}

function guessExt(url) {
  const ext = (String(url || '').split('.').pop() || '').split('?')[0].toLowerCase();
  if (ext === 'jpeg') return 'jpg';
  return ['jpg', 'png', 'webp', 'gif'].includes(ext) ? ext : 'jpg';
}

function getDownloadCode(product) {
  const itemId = String(product?.itemId || '').trim();
  if (/^\d+$/.test(itemId)) return itemId;
  return sanitizePathSegment(product?.title || 'item') || 'item';
}

async function setBackgroundStatus(message, type = '') {
  await storageSet({
    [BACKGROUND_STATUS_STORAGE]: {
      message,
      type,
      updatedAt: Date.now()
    }
  });
}

async function readJobQueue() {
  const stored = await storageGet(JOB_QUEUE_STORAGE);
  return Array.isArray(stored[JOB_QUEUE_STORAGE]) ? stored[JOB_QUEUE_STORAGE] : [];
}

async function writeJobQueue(queue) {
  await storageSet({ [JOB_QUEUE_STORAGE]: queue });
}

function createJobId() {
  return `job_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

async function enqueueJob(job) {
  const queue = await readJobQueue();
  const id = createJobId();
  queue.push({ ...job, id, queuedAt: Date.now() });
  await writeJobQueue(queue);
  chrome.alarms.create(JOB_ALARM_NAME, { when: Date.now() + 100 });
  return id;
}

async function upsertStoredProduct(product) {
  const stored = await storageGet(STORAGE_KEY);
  const products = Array.isArray(stored[STORAGE_KEY]) ? stored[STORAGE_KEY] : [];
  const index = products.findIndex(existing => existing.itemId === product.itemId);
  if (index >= 0) {
    products[index] = { ...products[index], ...product };
  } else {
    products.push(product);
  }
  await storageSet({ [STORAGE_KEY]: products });
  return index >= 0 ? 'updated' : 'added';
}

function extractItemIdFromUrl(url) {
  try {
    const itemId = new URL(url).searchParams.get('ItemId') || '';
    return /^\d+$/.test(itemId) ? itemId : null;
  } catch {
    return null;
  }
}

async function fetchAladinItem(ttbKey, itemId) {
  const url = `https://www.aladin.co.kr/ttb/api/ItemLookUp.aspx`
    + `?ttbkey=${encodeURIComponent(ttbKey)}`
    + `&itemIdType=ItemId`
    + `&ItemId=${encodeURIComponent(itemId)}`
    + `&output=js`
    + `&Version=20131101`
    + `&Cover=Big`
    + `&OptResult=authors,fulldescription`;

  const resp = await fetch(url, { cache: 'no-store' });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

  let text = (await resp.text()).trim();
  text = text.replace(/^\s*var\s+\w+\s*=\s*/, '').replace(/;\s*$/, '');
  const data = JSON.parse(text);
  if (!data.item || data.item.length === 0) return null;
  return data.item[0];
}

async function ensureScraperInjected(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    world: 'ISOLATED',
    files: ['content/aladin/content.js']
  });
}

async function scrapePageDetails(tabId, options) {
  await ensureScraperInjected(tabId);
  const results = await chrome.scripting.executeScript({
    target: { tabId },
    world: 'ISOLATED',
    func: scrapeOptions => {
      if (!globalThis.__ALADIN_SCRAPER__?.scrapePage) {
        throw new Error('スクレイパーを読み込めませんでした');
      }
      return globalThis.__ALADIN_SCRAPER__.scrapePage(scrapeOptions);
    },
    args: [options]
  });

  return results?.[0]?.result || null;
}

function buildProductRecord({ item, itemId, pageUrl, genre, scrapeResult }) {
  const apiDescription = String(item?.fulldescription || item?.description || '').trim();
  const pageDescription = String(scrapeResult?.description || '').trim();
  const magazineName = genre === '雑誌' ? extractMagazineName(item) : '';

  return {
    ...item,
    itemId,
    pageUrl,
    ジャンル: genre,
    magazineName,
    追加画像URL: scrapeResult?.additionalImagesJoined || '',
    additionalImages: scrapeResult?.additionalImages || [],
    description: pageDescription,
    apiDescription,
    pageDescription,
    basicInfo: scrapeResult?.basicInfo || ''
  };
}

async function processWriteCurrentPageJob(job) {
  const tabId = Number(job.tabId || 0);
  const tabUrl = String(job.tabUrl || '').trim();
  const ttbKey = String(job.ttbKey || '').trim();
  const gasUrl = String(job.gasUrl || '').trim();

  if (!tabId || !tabUrl) throw new Error('商品タブ情報を取得できません');
  if (!ttbKey) throw new Error('TTBキーを入力・保存してください');
  if (!gasUrl) throw new Error('GASのURLを保存してください');

  await setBackgroundStatus('API取得 + ページ解析中...', '');

  const itemId = extractItemIdFromUrl(tabUrl);
  if (!itemId) throw new Error('URLからItemIdを取得できません');

  const item = await fetchAladinItem(ttbKey, itemId);
  if (!item) throw new Error('商品が見つかりませんでした');

  const genre = detectGenre(item);
  await setBackgroundStatus(`API取得: ${genre} / ページ解析中...`, '');

  const scrapeResult = await scrapePageDetails(tabId, {
    genre,
    coverUrl: item.cover || ''
  });

  const product = buildProductRecord({
    item,
    itemId,
    pageUrl: tabUrl,
    genre,
    scrapeResult
  });

  const mode = await upsertStoredProduct(product);
  await setBackgroundStatus(`画像${product.additionalImages.length}枚取得 / ${genre} → 送信中...`, '');

  const response = await fetch(gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({
      action: 'updateAladinData',
      itemId: product.itemId,
      genre,
      title: product.title || '',
      magazineName: product.magazineName || '',
      author: product.author || '',
      publisher: product.publisher || '',
      pubDate: product.pubDate || '',
      priceSales: product.priceSales || '',
      cover: product.cover || '',
      isbn13: product.isbn13 || '',
      pageUrl: product.pageUrl || '',
      additionalImages: product.追加画像URL,
      description: product.description,
      basicInfo: product.basicInfo,
      categoryName: product.categoryName || '',
      mallType: product.mallType || ''
    })
  });

  if (!response.ok) {
    throw new Error(`GAS HTTP ${response.status}`);
  }

  const json = await response.json();
  if (!json.success) {
    throw new Error(json.error || '書き込み失敗');
  }

  const sheetActionLabel = json.mode === 'created'
    ? '追加'
    : json.mode === 'updated'
      ? '更新'
      : (mode === 'added' ? '追加' : '更新');
  const spreadsheetLabel = json.spreadsheetName || '保存先不明';
  const resolvedGenre = json.resolvedGenre || genre;
  const idLabel = json.spreadsheetId ? ` / ID:${String(json.spreadsheetId).slice(0, 8)}...` : ' / 旧GAS応答の可能性';
  await setBackgroundStatus(`✅ ${resolvedGenre} / シート${sheetActionLabel} / ${spreadsheetLabel} / ${json.sheet} 行${json.row}${idLabel}`, 'success');
}

async function processSaveProductsAssetsJob(job) {
  const products = Array.isArray(job.products) ? job.products : [];
  const mode = job.mode === 'images_only' ? 'images_only' : 'csv_and_images';
  const saveAs = job.saveAs !== false;
  if (!products.length) throw new Error('保存対象の商品がありません');

  await setBackgroundStatus(
    mode === 'images_only'
      ? '⏳ Downloads へ画像保存を開始しました'
      : '⏳ Downloads へ CSV+画像保存を開始しました',
    ''
  );

  let totalImages = 0;
  let successImages = 0;
  let csvCount = 0;

  for (const product of products) {
    const downloadCode = getDownloadCode(product);
    const rootedFolder = downloadCode;

    if (mode === 'csv_and_images') {
      const csvContent = buildCsvContent(product);
      const csvOk = await downloadCsvTask(csvContent, `${downloadCode}.csv`, saveAs);
      if (csvOk) csvCount += 1;
    }

    if (product.cover) {
      totalImages += 1;
      const ext = guessExt(product.cover);
      const result = await downloadImageTask(product.cover, `${rootedFolder}/1.${ext}`, saveAs);
      if (result.ok) successImages += 1;
    }

    const additionalUrls = String(product.追加画像URL || '')
      .split(';')
      .map(url => url.trim())
      .filter(Boolean);

    let sequence = 2;
    for (const url of additionalUrls) {
      totalImages += 1;
      const ext = guessExt(url);
      const result = await downloadImageTask(url, `${rootedFolder}/${sequence}.${ext}`, saveAs);
      if (result.ok) successImages += 1;
      sequence += 1;
    }
  }

  const prefix = mode === 'images_only'
    ? `✅ 画像保存 / ${products.length}商品 / ${successImages}/${totalImages}枚`
    : `✅ CSV${csvCount}/${products.length}件 / 画像${successImages}/${totalImages}枚`;
  await setBackgroundStatus(prefix, successImages < totalImages ? 'error' : 'success');

  return {
    mode,
    productCount: products.length,
    totalImages,
    successImages,
    failedImages: Math.max(0, totalImages - successImages),
    csvCount,
  };
}

async function processJobQueue() {
  if (isProcessingQueue) return;
  isProcessingQueue = true;

  try {
    while (true) {
      const queue = await readJobQueue();
      const job = queue[0];
      if (!job) break;

      try {
        if (job.type === 'writeCurrentPageToSheet') {
          await processWriteCurrentPageJob(job);
        } else if (job.type === 'saveProductsAssets') {
          await processSaveProductsAssetsJob(job);
        } else {
          await setBackgroundStatus(`❌ 未対応ジョブ: ${job.type}`, 'error');
        }
      } catch (error) {
        await setBackgroundStatus(`❌ ${error.message}`, 'error');
      }

      const currentQueue = await readJobQueue();
      await writeJobQueue(currentQueue.filter(item => item.id !== job.id));
    }
  } finally {
    isProcessingQueue = false;
  }
}

chrome.alarms.onAlarm.addListener(alarm => {
  if (alarm?.name === JOB_ALARM_NAME) {
    processJobQueue().catch(() => {});
  }
});

chrome.runtime.onStartup?.addListener(() => {
  processJobQueue().catch(() => {});
});

chrome.runtime.onInstalled?.addListener(() => {
  processJobQueue().catch(() => {});
});

chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request.action === 'downloadImage') {
    handleImageDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'downloadCsv') {
    handleCsvDownload(request, sendResponse);
    return true;
  }

  if (request.action === 'enqueueWriteCurrentPageToSheet') {
    enqueueJob({
      type: 'writeCurrentPageToSheet',
      tabId: request.tabId,
      tabUrl: request.tabUrl,
      ttbKey: request.ttbKey,
      gasUrl: request.gasUrl
    }).then(jobId => {
      setBackgroundStatus('⏳ バックグラウンドで書込を開始しました。タブ切替OKです', '').catch(() => {});
      sendResponse({ ok: true, jobId });
      chrome.alarms.create(JOB_ALARM_NAME, { when: Date.now() + 100 });
    }).catch(error => {
      sendResponse({ ok: false, error: error.message || 'enqueue failed' });
    });
    return true;
  }

  if (request.action === 'enqueueSaveProductsAssets') {
    processSaveProductsAssetsJob({
      products: Array.isArray(request.products) ? request.products : [],
      mode: request.mode,
      saveAs: request.saveAs !== false,
    }).then(result => {
      sendResponse({ ok: true, ...result });
    }).catch(error => {
      sendResponse({ ok: false, error: error.message || 'save failed' });
    });
    return true;
  }

  return false;
});







