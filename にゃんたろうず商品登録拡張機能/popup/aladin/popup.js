// popup.js - アラジン商品情報取得ツール

let products = [];
const STORAGE_KEY = 'aladin_products';
const TTB_KEY_STORAGE = 'aladin_ttb_key';
const GAS_URL_STORAGE = 'aladin_gas_url';
const BACKGROUND_STATUS_STORAGE = 'aladin_background_status';
const DEFAULT_TTB_KEY = 'ttbniyantarose1455001';
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbxZzoqvJRV9bY2VWXCcdQj0pRE-MPOIzu2d6lPSuMA9UVr9vp_iX_D5b2ENgNBXCcHYlw/exec';

function storageGet(keys) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(keys, result => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message || 'storage get failed'));
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
        reject(new Error(error.message || 'storage set failed'));
        return;
      }
      resolve();
    });
  });
}

function sendRuntimeMessage(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, response => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message || 'runtime message failed'));
        return;
      }
      resolve(response || {});
    });
  });
}

// ============================================================
// ジャンル判定（GASと同系統のロジック）
// ============================================================
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

function getGenreBadgeClass(genre) {
  if (genre === '音楽映像') return 'badge-media';
  if (genre === 'マンガ') return 'badge-comic';
  if (genre === 'グッズ') return 'badge-goods';
  return 'badge-book';
}

// ============================================================
// 初期化
// ============================================================
async function init() {
  const stored = await storageGet([STORAGE_KEY, TTB_KEY_STORAGE, GAS_URL_STORAGE, BACKGROUND_STATUS_STORAGE]);
  products = Array.isArray(stored[STORAGE_KEY]) ? stored[STORAGE_KEY] : [];

  const ttbKey = stored[TTB_KEY_STORAGE] || DEFAULT_TTB_KEY;
  const keyInput = document.getElementById('ttb-key-input');
  if (keyInput) keyInput.value = ttbKey;
  if (!stored[TTB_KEY_STORAGE]) {
    await storageSet({ [TTB_KEY_STORAGE]: DEFAULT_TTB_KEY });
  }

  const gasInput = document.getElementById('gas-url-input');
  const gasUrl = stored[GAS_URL_STORAGE] || DEFAULT_GAS_URL;
  if (gasInput) gasInput.value = gasUrl;
  if (!stored[GAS_URL_STORAGE]) {
    await storageSet({ [GAS_URL_STORAGE]: DEFAULT_GAS_URL });
  }
  renderList();
  updateUI();
  await refreshDirectSaveFolderStatus();

  const backgroundStatus = stored[BACKGROUND_STATUS_STORAGE];
  if (backgroundStatus?.message) {
    setStatus(backgroundStatus.message, backgroundStatus.type || '');
  }
}

init().catch(error => {
  setStatus(`❌ 初期化エラー: ${error.message}`, 'error');
});

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== 'local') return;

  if (changes[STORAGE_KEY]) {
    products = Array.isArray(changes[STORAGE_KEY].newValue) ? changes[STORAGE_KEY].newValue : [];
    renderList();
    updateUI();
    refreshDirectSaveFolderStatus().catch(() => {});
  }

  if (changes[BACKGROUND_STATUS_STORAGE]) {
    const status = changes[BACKGROUND_STATUS_STORAGE].newValue;
    if (status?.message) {
      setStatus(status.message, status.type || '');
    }
  }
});

// ============================================================
// UI
// ============================================================
function updateUI() {
  document.getElementById('item-count').textContent = `${products.length} 件`;
  document.getElementById('btn-download').disabled = products.length === 0;
  document.getElementById('btn-download-images').disabled = products.length === 0;
  document.getElementById('btn-clear').disabled = products.length === 0;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderList() {
  const container = document.getElementById('list-container');
  if (products.length === 0) {
    container.innerHTML = '<div class="empty-state">まだ商品が追加されていません</div>';
    return;
  }

  container.innerHTML = products.map((product, index) => {
    const title = escapeHtml(product.title || '（タイトルなし）');
    const genre = escapeHtml(product.ジャンル || '書籍');
    const price = escapeHtml(product.priceSales || '-');
    const itemId = escapeHtml(product.itemId || '');
    const cover = escapeHtml(product.cover || '');

    return `
      <div class="product-item">
        ${cover
          ? `<img class="product-img" src="${cover}" onerror="this.style.display='none'">`
          : '<div class="product-img"></div>'}
        <div class="product-info">
          <div class="product-name" title="${title}">${title}</div>
          <div class="product-meta">
            <span class="badge ${getGenreBadgeClass(product.ジャンル)}">${genre}</span>
            <span class="price">₩${price}</span>
            <span class="item-id">${itemId}</span>
          </div>
        </div>
        <button class="btn-remove" data-index="${index}" title="削除">×</button>
      </div>
    `;
  }).join('');

  container.querySelectorAll('.btn-remove').forEach(button => {
    button.addEventListener('click', () => {
      products.splice(Number.parseInt(button.dataset.index, 10), 1);
      save();
      renderList();
      updateUI();
      refreshDirectSaveFolderStatus().catch(() => {});
    });
  });
}

function setStatus(message, type = '') {
  const element = document.getElementById('status');
  element.textContent = message;
  element.className = `status-bar ${type}`;
}

function save() {
  storageSet({ [STORAGE_KEY]: products }).catch(() => {});
}

function isDirectFolderSaveAvailable() {
  return typeof window.showDirectoryPicker === 'function'
    && typeof pickAndSaveFolder === 'function'
    && typeof getSavedFolderHandle === 'function'
    && typeof getSavedFolderName === 'function'
    && typeof ensureFolderPermission === 'function'
    && typeof saveTextFileToFolder === 'function'
    && typeof saveImagesToFolder === 'function';
}

async function refreshDirectSaveFolderStatus() {
  const input = document.getElementById('folder-status');
  if (!input) return;

  if (!isDirectFolderSaveAvailable()) {
    input.value = 'このブラウザでは直接保存を使えません';
    return;
  }

  const folderName = await getSavedFolderName().catch(() => '');
  input.value = folderName || '未選択';
}

function normalizeDirectSaveError(error) {
  if (error?.name === 'AbortError') {
    return new Error('保存先フォルダの選択をキャンセルしました');
  }
  return error instanceof Error
    ? error
    : new Error(String(error || '保存先フォルダを準備できませんでした'));
}

async function prepareDirectSaveContext(options = {}) {
  if (!isDirectFolderSaveAvailable()) {
    return null;
  }

  const promptIfMissing = options.promptIfMissing === true;
  let rootHandle = await getSavedFolderHandle().catch(() => null);
  let folderName = await getSavedFolderName().catch(() => '');

  if (!promptIfMissing) {
    if (!rootHandle && !folderName) {
      return null;
    }
    return {
      rootHandle: null,
      folderName: folderName || rootHandle?.name || '',
    };
  }

  if (rootHandle) {
    const granted = await ensureFolderPermission(rootHandle).catch(() => false);
    if (granted) {
      return {
        rootHandle,
        folderName: folderName || rootHandle.name || '',
      };
    }
  }

  try {
    folderName = await pickAndSaveFolder();
  } catch (error) {
    throw normalizeDirectSaveError(error);
  }

  rootHandle = await getSavedFolderHandle().catch(() => null);
  if (!rootHandle) {
    throw new Error('保存先フォルダを取得できませんでした');
  }

  const granted = await ensureFolderPermission(rootHandle).catch(() => false);
  if (!granted) {
    throw new Error('保存先フォルダへの書き込み権限がありません');
  }

  await refreshDirectSaveFolderStatus();
  return {
    rootHandle,
    folderName: folderName || rootHandle.name || '',
  };
}

async function chooseDirectSaveFolder() {
  const directSaveContext = await prepareDirectSaveContext({ promptIfMissing: true });
  await refreshDirectSaveFolderStatus();
  return directSaveContext?.folderName || '';
}

function describeSaveDestination(directSaveContext, fallbackLabel = 'Downloads') {
  if (directSaveContext?.rootHandle) {
    return `直接保存:${directSaveContext.folderName || '選択フォルダ'}`;
  }
  return fallbackLabel;
}

function sanitizeDownloadSegment(segment) {
  return String(segment || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '') || 'item';
}

function buildCsvContentForDirectSave(product) {
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

function guessDirectImageExt(url) {
  const ext = (String(url || '').split('.').pop() || '').split('?')[0].toLowerCase();
  if (ext === 'jpeg') return 'jpg';
  return ['jpg', 'png', 'webp', 'gif'].includes(ext) ? ext : 'jpg';
}

function getDirectSaveDownloadCode(product) {
  const itemId = String(product?.itemId || '').trim();
  if (/^\d+$/.test(itemId)) return itemId;
  return sanitizeDownloadSegment(product?.title || 'item');
}

function collectDirectSaveImageUrls(product) {
  const urls = [
    String(product?.cover || '').trim(),
    ...String(product?.追加画像URL || '')
      .split(';')
      .map(url => url.trim())
      .filter(Boolean),
  ].filter(url => /^https?:\/\//i.test(url));

  return [...new Set(urls)];
}

async function saveProductsDirect(productsToSave, mode, directSaveContext) {
  const products = Array.isArray(productsToSave) ? productsToSave : [];
  if (!directSaveContext?.rootHandle) {
    throw new Error('保存先フォルダを準備できませんでした');
  }

  let totalImages = 0;
  let successImages = 0;
  let csvCount = 0;

  for (const product of products) {
    const downloadCode = getDirectSaveDownloadCode(product);
    const rootedFolder = downloadCode;

    if (mode === 'csv_and_images') {
      const csvContent = buildCsvContentForDirectSave(product);
      await saveTextFileToFolder(directSaveContext.rootHandle, `${downloadCode}.csv`, csvContent);
      csvCount += 1;
    }

    const imageUrls = collectDirectSaveImageUrls(product);
    if (imageUrls.length > 0) {
      const saveResult = await saveImagesToFolder(directSaveContext.rootHandle, rootedFolder, imageUrls, null, {
        forceJpeg: false,
      });
      totalImages += saveResult.total;
      successImages += saveResult.success;
    }
  }

  return {
    mode,
    productCount: products.length,
    totalImages,
    successImages,
    failedImages: Math.max(0, totalImages - successImages),
    csvCount,
    folderName: directSaveContext.folderName || '',
    imageExtensionMode: 'keep-original',
  };
}

// ============================================================
// 商品収集
// ============================================================
function extractItemIdFromUrl(url) {
  try {
    const itemId = new URL(url).searchParams.get('ItemId') || '';
    return /^\d+$/.test(itemId) ? itemId : null;
  } catch (_error) {
    return null;
  }
}

async function getActiveAladinTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !tab.url) {
    throw new Error('アクティブなタブを取得できません');
  }
  if (!/^https:\/\/www\.aladin\.co\.kr\/shop\/wproduct\.aspx/i.test(tab.url)) {
    throw new Error('アラジンの商品ページを開いてください');
  }
  return tab;
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
    apiDescription: String(item?.fulldescription || item?.description || '').trim(),
    pageDescription,
    basicInfo: scrapeResult?.basicInfo || ''
  };
}

function upsertProduct(product) {
  const index = products.findIndex(existing => existing.itemId === product.itemId);
  if (index >= 0) {
    products[index] = { ...products[index], ...product };
    return 'updated';
  }
  products.push(product);
  return 'added';
}

async function collectCurrentPageProduct() {
  const stored = await storageGet(TTB_KEY_STORAGE);
  const ttbKey = String(stored[TTB_KEY_STORAGE] || '').trim();
  if (!ttbKey) {
    throw new Error('TTBキーを入力・保存してください');
  }

  const tab = await getActiveAladinTab();
  const itemId = extractItemIdFromUrl(tab.url);
  if (!itemId) {
    throw new Error('URLからItemIdを取得できません');
  }

  const item = await fetchAladinItem(ttbKey, itemId);
  if (!item) {
    throw new Error('商品が見つかりませんでした');
  }

  const genre = detectGenre(item);
  setStatus(`API取得: ${genre} / ページ解析中...`, '');
  const scrapeResult = await scrapePageDetails(tab.id, {
    genre,
    coverUrl: item.cover || ''
  });

  return {
    tab,
    itemId,
    item,
    genre,
    scrapeResult,
    product: buildProductRecord({
      item,
      itemId,
      pageUrl: tab.url,
      genre,
      scrapeResult
    })
  };
}

// ============================================================
// 保存ボタン類
// ============================================================
document.getElementById('btn-save-key')?.addEventListener('click', async () => {
  const key = document.getElementById('ttb-key-input')?.value?.trim() || '';
  if (!key) {
    setStatus('⚠ TTBキーを入力してください', 'error');
    return;
  }
  await storageSet({ [TTB_KEY_STORAGE]: key });
  setStatus('✅ TTBキーを保存しました', 'success');
});

document.getElementById('btn-save-gas-url')?.addEventListener('click', async () => {
  const url = document.getElementById('gas-url-input')?.value?.trim() || '';
  if (!url) {
    setStatus('⚠ GASのURLを入力してください', 'error');
    return;
  }
  await storageSet({ [GAS_URL_STORAGE]: url });
  setStatus('✅ GAS URLを保存しました', 'success');
});

// ============================================================
// 追加ボタン
// ============================================================
document.getElementById('btn-add').addEventListener('click', async () => {
  setStatus('API取得 + ページ解析中...', '');

  try {
    const { product } = await collectCurrentPageProduct();
    upsertProduct(product);
    save();
    renderList();
  updateUI();
  await refreshDirectSaveFolderStatus();
    setStatus(`✅ ${product.ジャンル} として追加`, 'success');
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

// ============================================================
// ダウンロードボタン
// ============================================================
document.getElementById('btn-download').addEventListener('click', async () => {
  if (products.length === 0) return;
  setStatus('⏳ 保存先を確認中...', '');

  try {
    const directSaveContext = await prepareDirectSaveContext({ promptIfMissing: true });
    setStatus('⏳ CSV+画像を保存中...', '');

    const result = directSaveContext?.rootHandle
      ? await saveProductsDirect(products, 'csv_and_images', directSaveContext)
      : await sendRuntimeMessage({
          action: 'enqueueSaveProductsAssets',
          products,
          mode: 'csv_and_images',
          saveAs: true,
        });
    if (!directSaveContext?.rootHandle && !result?.ok) {
      throw new Error(result?.error || 'CSV+画像保存に失敗しました');
    }

    const destinationLabel = describeSaveDestination(directSaveContext, 'Downloads');
    setStatus(
      `✅ ${destinationLabel} に CSV:${result.csvCount}/${products.length}件 / 画像:${result.successImages}/${result.totalImages}枚保存 / 商品フォルダ:${result.productCount}` +
        (result.failedImages > 0 ? `（失敗:${result.failedImages}）` : ''),
      result.failedImages > 0 ? 'error' : 'success'
    );
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-download-images').addEventListener('click', async () => {
  if (products.length === 0) return;
  setStatus('⏳ 保存先を確認中...', '');

  try {
    const directSaveContext = await prepareDirectSaveContext({ promptIfMissing: true });
    setStatus('⏳ 画像を保存中...', '');

    const result = directSaveContext?.rootHandle
      ? await saveProductsDirect(products, 'images_only', directSaveContext)
      : await sendRuntimeMessage({
          action: 'enqueueSaveProductsAssets',
          products,
          mode: 'images_only',
          saveAs: true,
        });
    if (!directSaveContext?.rootHandle && !result?.ok) {
      throw new Error(result?.error || '画像保存に失敗しました');
    }

    const destinationLabel = describeSaveDestination(directSaveContext, 'Downloads');
    setStatus(
      `✅ ${destinationLabel} に画像:${result.successImages}/${result.totalImages}枚保存 / 商品フォルダ:${result.productCount}` +
        (result.failedImages > 0 ? `（失敗:${result.failedImages}）` : ''),
      result.failedImages > 0 ? 'error' : 'success'
    );
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

// ============================================================
// クリアボタン
// ============================================================
document.getElementById('btn-clear').addEventListener('click', () => {
  if (products.length === 0) return;
  if (!confirm(`${products.length}件のリストをすべて削除しますか？`)) return;
  products = [];
  save();
  renderList();
  updateUI();
  refreshDirectSaveFolderStatus().catch(() => {});
  setStatus('リストをクリアしました', '');
});

// ============================================================
// GAS連携機能
// ============================================================
document.getElementById('btn-send-to-sheet')?.addEventListener('click', async () => {
  const stored = await storageGet([GAS_URL_STORAGE, TTB_KEY_STORAGE]);
  const gasUrl = String(stored[GAS_URL_STORAGE] || '').trim();
  const ttbKey = String(stored[TTB_KEY_STORAGE] || '').trim();
  if (!gasUrl) {
    setStatus('⚠ GASのURLを保存してください', 'error');
    return;
  }
  if (!ttbKey) {
    setStatus('⚠ TTBキーを入力・保存してください', 'error');
    return;
  }

  setStatus('⏳ バックグラウンドで書込を開始しています...', '');

  try {
    const tab = await getActiveAladinTab();
    const response = await sendRuntimeMessage({
      action: 'enqueueWriteCurrentPageToSheet',
      tabId: tab.id,
      tabUrl: tab.url,
      ttbKey,
      gasUrl
    });

    if (!response?.ok) {
      throw new Error(response?.error || '書込ジョブを開始できませんでした');
    }

    setStatus('⏳ バックグラウンドで書込開始。タブ切替OKです', '');
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});















