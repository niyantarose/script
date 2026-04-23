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

function detectKoreanBookCategory(source) {
  const categoryName = String(source?.categoryName || '').trim();
  const title = String(source?.title || '').trim();
  const basicInfo = String(source?.basicInfo || '').trim();
  const description = String(source?.description || source?.fulldescription || '').trim();
  const mainText = [categoryName, title, basicInfo].join(' ').toLowerCase();
  const combined = [mainText, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(combined)) return 'OST';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/essay|에세이|산문/.test(mainText)) return 'エッセイ';
  if (/참고서|수험서|문제집|기출|모의고사|수능|내신|검정고시|자격증|공무원|고시|임용|편입|leet|meet|deet|psat|ncs|toeic|toefl|ielts|teps|jlpt|hsk|topik|参考書|問題集|過去問|受験|資格/.test(mainText)) return '参考書';
  if (/교재|학습지|학습서|워크북|work\s*book|workbook|text\s*book|textbook|student\s*book|activity\s*book|teacher'?s\s*book|course\s*book|教材|教科書/.test(mainText)) return '教材';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북|예체능|미술|디자인|사진/.test(combined)) return 'アートブック';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  if (/novel|소설|라이트노벨|ライトノベル|light\s*novel|문학|시집/.test(combined)) return '小説';
  return '';
}

function detectKoreanMangaCategory(source) {
  const categoryName = String(source?.categoryName || '').trim();
  const title = String(source?.title || '').trim();
  const subtitle = String(source?.subInfo?.subTitle || '').trim();
  const combined = [categoryName, title, subtitle].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/essay|에세이/.test(combined)) return 'エッセイ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북/.test(combined)) return 'アートブック';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/novel|소설|라이트노벨|ライトノベル|light\s*novel/.test(combined)) return '小説';
  return 'まんが';
}

function detectKoreanMediaCategory(source) {
  const categoryName = String(source?.categoryName || '').trim();
  const title = String(source?.title || '').trim();
  const basicInfo = String(source?.basicInfo || '').trim();
  const description = String(source?.description || '').trim();
  const titleAndCategory = [title, categoryName].join(' ').toLowerCase();
  const combined = [title, categoryName, basicInfo, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/\bdvd\b|\[dvd\]|dvd\//.test(titleAndCategory)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(titleAndCategory)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(titleAndCategory)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(titleAndCategory)) return 'OST';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(titleAndCategory)) return 'CD';
  if (/\bdvd\b|\[dvd\]|dvd\//.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(combined)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(combined)) return 'OST';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(combined)) return 'CD';
  return '';
}

function detectSheetCategory(genre, item, scrapeResult = {}) {
  const source = {
    ...item,
    basicInfo: scrapeResult?.basicInfo || '',
    description: scrapeResult?.description || item?.fulldescription || item?.description || '',
  };
  if (genre === '書籍') return detectKoreanBookCategory(source);
  if (genre === 'マンガ') return detectKoreanMangaCategory(source);
  if (genre === '音楽映像') return detectKoreanMediaCategory(source);
  if (genre === '雑誌') return '雑誌';
  if (genre === 'グッズ') return 'グッズ';
  return '';
}

function getProductSheetCategory(product) {
  return product?.sheetCategory
    || product?.登録カテゴリ
    || detectSheetCategory(product?.ジャンル || detectGenre(product), product, product);
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
    const sheetCategory = escapeHtml(getProductSheetCategory(product));
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
            ${sheetCategory ? `<span class="badge badge-category">${sheetCategory}</span>` : ''}
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

function sanitizeDownloadSegment(segment) {
  return String(segment || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '') || 'item';
}

function buildAladinCsvContent(product) {
  const toCsv = (headers, row) =>
    '\uFEFF' + [headers, row]
      .map(record => record.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\r\n');

  const genre = product.ジャンル;
  const additional = product.追加画像URL || '';
  const description = product.description || '';
  const sheetCategory = getProductSheetCategory(product) || product.categoryName || '';

  if (genre === '音楽映像') {
    return toCsv(
      ['商品名（原題）', 'アーティスト', 'レーベル', '発売日', '原価', 'ジャンル分類', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, product.author, product.publisher, product.pubDate,
       product.priceSales, sheetCategory, product.cover, additional,
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
       product.priceSales, sheetCategory, product.isbn13,
       product.cover, additional, description, product.pageUrl]
    );
  }
  if (genre === 'グッズ') {
    return toCsv(
      ['商品名（原題）', '発売日', '原価', 'ジャンル分類', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, product.pubDate, product.priceSales, sheetCategory,
       product.cover, additional, description, product.pageUrl]
    );
  }

  return toCsv(
    ['商品名（原題）', '著者', '出版社', '発売日', '原価', 'ジャンル分類', 'ISBN', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
    [product.title, product.author, product.publisher, product.pubDate,
     product.priceSales, sheetCategory, product.isbn13,
     product.cover, additional, description, product.pageUrl]
  );
}

function guessAladinImageExt(url) {
  const ext = (String(url || '').split('.').pop() || '').split('?')[0].toLowerCase();
  if (ext === 'jpeg') return 'jpg';
  return ['jpg', 'png', 'webp', 'gif'].includes(ext) ? ext : 'jpg';
}

function getAladinDownloadCode(product) {
  const itemId = String(product?.itemId || '').trim();
  if (/^\d+$/.test(itemId)) return itemId;
  return sanitizeDownloadSegment(product?.title || 'item');
}

function collectAladinImageUrls(product) {
  const urls = [
    String(product?.cover || '').trim(),
    ...String(product?.追加画像URL || '')
      .split(';')
      .map(url => url.trim())
      .filter(Boolean),
  ].filter(url => /^https?:\/\//i.test(url));

  return [...new Set(urls)];
}

function triggerAladinAnchorDownload(downloadName, url) {
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

async function downloadPreparedCsv(csvText, filename) {
  const response = await sendRuntimeMessage({ action: 'prepareAladinCsvDownload', csvText, filename });
  if (!response?.ok || !response.url || !response.downloadName) {
    return false;
  }
  triggerAladinAnchorDownload(response.downloadName, response.url);
  await waitForQueuedDownloadTick();
  return true;
}

async function downloadPreparedImage(imageUrl, filename) {
  const response = await sendRuntimeMessage({ action: 'prepareAladinImageDownload', url: imageUrl, filename });
  if (!response?.ok || !response.url || !response.downloadName) {
    return false;
  }
  triggerAladinAnchorDownload(response.downloadName, response.url);
  await waitForQueuedDownloadTick();
  return true;
}

async function saveProductsWithDownloadFlow(productsToSave, mode) {
  const products = Array.isArray(productsToSave) ? productsToSave : [];
  let totalImages = 0;
  let successImages = 0;
  let csvCount = 0;

  for (const product of products) {
    const downloadCode = getAladinDownloadCode(product);

    if (mode === 'csv_and_images') {
      const csvContent = buildAladinCsvContent(product);
      const csvOk = await downloadPreparedCsv(csvContent, `${downloadCode}.csv`);
      if (csvOk) csvCount += 1;
    }

    const imageUrls = collectAladinImageUrls(product);
    let sequence = 1;
    for (const imageUrl of imageUrls) {
      totalImages += 1;
      const ext = guessAladinImageExt(imageUrl);
      const ok = await downloadPreparedImage(imageUrl, `${downloadCode}/${sequence}.${ext}`);
      if (ok) successImages += 1;
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
  const sheetCategory = detectSheetCategory(genre, item, {
    ...scrapeResult,
    description: pageDescription,
  });

  return {
    ...item,
    itemId,
    pageUrl,
    ジャンル: genre,
    登録カテゴリ: sheetCategory,
    sheetCategory,
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
  setStatus('⏳ CSV+画像を出力中...', '');

  try {
    const result = await saveProductsWithDownloadFlow(products, 'csv_and_images');
    const destinationLabel = '手動保存ダイアログ';
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
  setStatus('⏳ 画像を出力中...', '');

  try {
    const result = await saveProductsWithDownloadFlow(products, 'images_only');
    const destinationLabel = '手動保存ダイアログ';
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
  if (products.length === 0) {
    setStatus('⚠ 先に商品を追加してください', 'error');
    return;
  }

  setStatus(`⏳ バックグラウンドで${products.length}件の書込を開始しています...`, '');

  try {
    const response = await sendRuntimeMessage({
      action: 'enqueueWriteProductsToSheet',
      products,
      gasUrl,
      ttbKey
    });

    if (!response?.ok) {
      throw new Error(response?.error || '書込ジョブを開始できませんでした');
    }

    setStatus(`⏳ ${products.length}件のシート書込開始。タブ切替OKです`, '');
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});















