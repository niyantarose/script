// background.js - アラジン拡張
// ダウンロードとシート書込をバックグラウンドで継続実行する

const STORAGE_KEY = 'aladin_products';
const BACKGROUND_STATUS_STORAGE = 'aladin_background_status';
const JOB_QUEUE_STORAGE = 'aladin_job_queue';
const JOB_ALARM_NAME = 'aladin-job-queue';
const ALADIN_DOWNLOAD_MARKER = 'nyantarose_aladin___-_';
const pendingAladinFilenameByMarker = new Map();
const ALADIN_FILENAME_HOOK_TTL_MS = 30000;

let isProcessingQueue = false;
let aladinFilenameHookCleanupTimer = null;
let aladinFilenameHookRegistered = false;


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

function getPendingAladinFilenameValue(entry) {
  if (!entry) return '';
  if (typeof entry === 'string') return entry;
  return String(entry.filename || '').trim();
}

function prunePendingAladinFilenameMarkers(now = Date.now()) {
  for (const [markerFilename, entry] of pendingAladinFilenameByMarker.entries()) {
    const createdAt = Number(entry?.createdAt || 0);
    if (createdAt && now - createdAt > ALADIN_FILENAME_HOOK_TTL_MS) {
      pendingAladinFilenameByMarker.delete(markerFilename);
    }
  }
}

function releaseAladinFilenameHookIfIdle() {
  prunePendingAladinFilenameMarkers();
  if (pendingAladinFilenameByMarker.size > 0) return;
  if (aladinFilenameHookCleanupTimer) {
    clearTimeout(aladinFilenameHookCleanupTimer);
    aladinFilenameHookCleanupTimer = null;
  }
  if (aladinFilenameHookRegistered) {
    chrome.downloads.onDeterminingFilename.removeListener(handleAladinDeterminingFilename);
    aladinFilenameHookRegistered = false;
  }
}

function scheduleAladinFilenameHookCleanup() {
  if (aladinFilenameHookCleanupTimer) {
    clearTimeout(aladinFilenameHookCleanupTimer);
  }
  aladinFilenameHookCleanupTimer = setTimeout(() => {
    releaseAladinFilenameHookIfIdle();
  }, ALADIN_FILENAME_HOOK_TTL_MS);
}

function ensureAladinFilenameHook() {
  if (!aladinFilenameHookRegistered) {
    chrome.downloads.onDeterminingFilename.addListener(handleAladinDeterminingFilename);
    aladinFilenameHookRegistered = true;
  }
  scheduleAladinFilenameHookCleanup();
}

function buildAladinMarkerFilename(targetFilename = 'download.bin') {
  const safeTargetFilename = sanitizeRelativePath(targetFilename, 'download.bin');
  const extensionMatch = safeTargetFilename.match(/(\.[a-z0-9]+)$/i);
  const extension = extensionMatch ? extensionMatch[1] : '.bin';
  const token = `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const markerFilename = `${ALADIN_DOWNLOAD_MARKER}${token}${extension}`;
  pendingAladinFilenameByMarker.set(markerFilename, {
    filename: safeTargetFilename,
    createdAt: Date.now(),
  });
  ensureAladinFilenameHook();
  return markerFilename;
}

function consumePendingAladinFilenameByMarker(itemFilename = '') {
  const normalized = String(itemFilename || '')
    .replace(/\\/g, '/')
    .split('/')
    .map(part => String(part || '').trim())
    .filter(Boolean)
    .pop() || '';
  if (!normalized || !normalized.includes(ALADIN_DOWNLOAD_MARKER)) return '';

  const direct = pendingAladinFilenameByMarker.get(normalized);
  if (direct) {
    pendingAladinFilenameByMarker.delete(normalized);
    return getPendingAladinFilenameValue(direct);
  }

  for (const [markerFilename, entry] of pendingAladinFilenameByMarker.entries()) {
    if (normalized.includes(markerFilename)) {
      pendingAladinFilenameByMarker.delete(markerFilename);
      return getPendingAladinFilenameValue(entry);
    }
  }

  return '';
}

function handleAladinDeterminingFilename(item, suggest) {
  prunePendingAladinFilenameMarkers();
  const desiredFilename = consumePendingAladinFilenameByMarker(item?.filename || '');
  if (desiredFilename) {
    suggest({ filename: desiredFilename, conflictAction: 'uniquify' });
  } else {
    suggest();
  }
  releaseAladinFilenameHookIfIdle();
}

async function fetchAladinAssetAsDataUrl(url) {
  const response = await fetch(url, {
    credentials: 'omit',
    cache: 'no-store',
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${url}`);
  }
  const blob = await response.blob();
  if (!blob || !blob.size) {
    throw new Error(`empty blob: ${url}`);
  }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const mimeType = blob.type || 'application/octet-stream';
  return `data:${mimeType};base64,${bytesToBase64(bytes)}`;
}

async function prepareAladinImageDownload(request) {
  const url = String(request.url || '').trim();
  const filename = sanitizeRelativePath(request.filename, '1.jpg');
  if (!url) {
    throw new Error('url missing');
  }
  const dataUrl = await fetchAladinAssetAsDataUrl(url);
  const downloadName = buildAladinMarkerFilename(filename);
  return { ok: true, url: dataUrl, downloadName, filename };
}

async function prepareAladinCsvDownload(request) {
  const csvText = String(request.csvText || '');
  const filename = sanitizeRelativePath(request.filename, 'export.csv');
  if (!csvText) {
    throw new Error('csv text missing');
  }
  const csvBytes = new TextEncoder().encode(csvText);
  const dataUrl = `data:text/csv;charset=utf-8;base64,${bytesToBase64(csvBytes)}`;
  const downloadName = buildAladinMarkerFilename(filename);
  return { ok: true, url: dataUrl, downloadName, filename };
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
  if (/novel|소설|라이트노ベル|ライトノベル|light\s*novel/.test(combined)) return '小説';
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

function normalizeAladinWorksTitle(product) {
  const originalTitle = String(product?.title || '').trim();
  if (!originalTitle) return '';

  let title = originalTitle
    .replace(/\u3000/g, ' ')
    .replace(/[（）]/g, match => match === '（' ? '(' : ')')
    .replace(/[［］]/g, match => match === '［' ? '[' : ']')
    .replace(/[：]/g, ':')
    .replace(/[‐−–—―]/g, '-')
    .replace(/\bo\.?\s*s\.?\s*t\.?\b/gi, 'OST')
    .replace(/\s+/g, ' ')
    .trim();

  const category = getProductSheetCategory(product);
  const isMedia = product?.ジャンル === '音楽映像' || ['CD', 'LP', 'OST', 'DVD', 'Blu-ray'].includes(category);
  if (!isMedia) return title;

  title = title.replace(/^\s*\[[^\]]*(?:4k|blu[-\s]?ray|블루레이|dvd|uhd|ubd)[^\]]*\]\s*/i, '');

  const dashMatch = title.match(/^(.{1,45}?)\s+-\s+(.+)$/);
  if (dashMatch) {
    const beforeDash = dashMatch[1].trim();
    const afterDash = dashMatch[2].trim();
    const afterStartsWithAlbumTitle =
      /^(?:정규|미니|싱글|스페셜|리패키지)\s*\d+\s*(?:집|앨범)\b/i.test(afterDash)
      || /^\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\b/i.test(afterDash)
      || /(?:no tragedy|revive)/i.test(afterDash);
    const afterLooksLikeOnlyPackage =
      /(?:상자|박스|부클릿|북클릿|소책자|컵받침대|거치대|종이\s*거치대|파우치|포토카드|포토북|카드|커버|세트|랜덤|booklet|photocard|poster|pouch|nfc|\d+\s*(?:종|p|disc)|cd\s*\()/i.test(afterDash);

    if (afterStartsWithAlbumTitle) {
      title = afterDash;
    } else if (beforeDash && afterLooksLikeOnlyPackage) {
      title = beforeDash;
    }
  }
  title = title
    .replace(/^\s*(?:정규|미니|싱글|스페셜|리패키지)\s*\d+\s*(?:집|앨범)\s*/i, '')
    .replace(/^\s*\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\s*/i, '')
    .replace(/\[[^\]]*(?:180g|white\s*marble|vinyl|lp|ver\.?|한정|限定|picture|픽처|photocard|poster|cd|dvd|blu[-\s]?ray|ubd|uhd|disc)[^\]]*\]/gi, ' ')
    .replace(/\([^)]*(?:album\s*ver\.?|mubeat|pouch|파우치|mini\s*cd|미니\s*cd|photocard|포토카드|포토북|booklet|소책자|nfc|disc|cd|dvd|blu[-\s]?ray|ubd|uhd|\d+\s*종|\d+\s*p)[^)]*\)/gi, ' ')
    .replace(/\s+-\s+.*(?:커버|바이닐|포토|파우치|소책자|부클릿|북클릿|상자|박스|컵받침대|거치대|종이\s*거치대|랜덤|booklet|poster|photocard|album\s*ver\.?|nfc|카드|세트|컵).*$/i, ' ')
    .replace(/\s*:\s*(?:스틸북|풀슬립|full\s*slip|steelbook).*$/i, ' ')
    .replace(/\b(?:4k\s*uhd|4k|uhd|ubd|2d|blu[-\s]?ray|dvd|cd|lp|vinyl|record)\b/gi, ' ')
    .replace(/(?:블루레이|스틸북|풀슬립|소책자|포토카드|포토북|북클릿|파우치|한정반|한정판|게이트폴드\s*커버|바이닐)/gi, ' ')
    .replace(/\b(?:a|b|c|d)\s*ver\.?\b/gi, ' ')
    .replace(/\b(?:album|mubeat\s*album)\s*ver\.?\b/gi, ' ')
    .replace(/\d+\s*(?:disc|종|p)\b/gi, ' ')
    .replace(/\s*(?:OST|O\.S\.T\.?|original\s*sound\s*track|사운드트랙)\s*$/i, '')
    .replace(/[<＞>]+/g, ' ')
    .replace(/[『』「」"'`]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!title) {
    const subTitle = String(product?.subInfo?.subTitle || '').trim();
    if (subTitle) {
      const cleaned = subTitle
        .replace(/　/g, ' ')
        .replace(/^\s*(?:정규|미니|싱글|스페셜|리패키지)\s*\d*\s*(?:집|앨범)?\s*$/i, '')
        .replace(/^\s*\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\s*$/i, '')
        .replace(/\s+/g, ' ')
        .trim();
      if (cleaned) return cleaned;
    }
    return '';
  }

  title = title
    .replace(/^(?:넷플릭스\s*)?시리즈\s+(.+?)\s+OST$/i, '$1 OST')
    .replace(/\bost\b/gi, 'OST');
  return title;
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
  const sheetCategory = getProductSheetCategory(product) || product.categoryName || '';

  if (genre === '音楽映像') {
    const worksTitle = product.worksTitle || product.productTitle || normalizeAladinWorksTitle(product);
    return toCsv(
      ['商品名（原題）', '商品名(タイトル)', 'アーティスト', 'レーベル', '発売日', '原価', 'ジャンル分類', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジンURL'],
      [product.title, worksTitle, product.author, product.publisher, product.pubDate,
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
  const sheetCategory = detectSheetCategory(genre, item, {
    ...scrapeResult,
    description: pageDescription || apiDescription,
  });
  const productBase = {
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
    apiDescription,
    pageDescription,
    basicInfo: scrapeResult?.basicInfo || ''
  };
  const worksTitle = normalizeAladinWorksTitle(productBase);

  return {
    ...productBase,
    worksTitle,
    productTitle: worksTitle
  };
}

function buildSheetPayload(product, fallbackGenre) {
  const genre = fallbackGenre || product.ジャンル || detectGenre(product);
  const rawItem = {
    source: 'aladin',
    itemId: product.itemId,
    genre,
    title: product.title || '',
    worksTitle: product.worksTitle || product.productTitle || normalizeAladinWorksTitle(product),
    productTitle: product.productTitle || product.worksTitle || normalizeAladinWorksTitle(product),
    magazineName: product.magazineName || '',
    author: product.author || '',
    publisher: product.publisher || '',
    pubDate: product.pubDate || '',
    priceSales: product.priceSales || '',
    cover: product.cover || '',
    isbn13: product.isbn13 || '',
    pageUrl: product.pageUrl || '',
    additionalImages: product.追加画像URL || '',
    description: product.description || product.pageDescription || '',
    basicInfo: product.basicInfo || '',
    categoryName: product.categoryName || '',
    sheetCategory: getProductSheetCategory(product) || '',
    normalizedCategory: getProductSheetCategory(product) || '',
    mallType: product.mallType || ''
  };
  const titleAnalysis = product.titleAnalysis || (typeof globalThis.analyzeProductTitle === 'function'
    ? globalThis.analyzeProductTitle(rawItem)
    : null);
  return {
    action: 'upsertProductWithLookup',
    source: 'aladin',
    rawItem,
    titleAnalysis,
    japaneseTitleLookup: product.japaneseTitleLookup || null,
    ...rawItem
  };
}

function buildLegacySheetPayload(product, fallbackGenre) {
  const payload = buildSheetPayload(product, fallbackGenre);
  return {
    ...payload.rawItem,
    action: 'updateAladinData',
    titleAnalysis: payload.titleAnalysis,
    japaneseTitleLookup: payload.japaneseTitleLookup || null
  };
}

async function postAladinPayloadToGas(gasUrl, payload) {
  const response = await fetch(gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error(`GAS HTTP ${response.status}`);
  }

  const json = await response.json();
  return json;
}

function shouldFallbackToLegacyAladinAction(json) {
  const message = String(json?.error || '').toLowerCase();
  return json && json.success === false && /unknown action|unsupported action/.test(message);
}

async function lookupJapaneseTitleForProduct(gasUrl, product, fallbackGenre) {
  const payload = buildSheetPayload(product, fallbackGenre);
  if (!payload.titleAnalysis) return product;
  try {
    let json = await postAladinPayloadToGas(gasUrl, {
      action: 'lookupJapaneseTitle',
      source: 'aladin',
      rawItem: payload.rawItem,
      titleAnalysis: payload.titleAnalysis,
    });
    const lu = json && json.lookup ? json.lookup : null;
    if (
      lu
      && !String(lu.japaneseTitle || '').trim()
      && globalThis.titleLookupMangaUpdates
      && typeof globalThis.titleLookupMangaUpdates.lookupJapaneseTitle === 'function'
    ) {
      const ta = payload.titleAnalysis;
      const types = new Set(['manga', 'bl_manga', 'goods', 'light_novel', 'unknown']);
      if (types.has(String(ta && ta.itemType || '').trim())) {
        const q = String((ta && (ta.normalizedSearchTitle || ta.extractedWorkTitle)) || '').trim();
        if (q) {
          try {
            const jp = String(await globalThis.titleLookupMangaUpdates.lookupJapaneseTitle(q)).trim();
            if (jp) {
              json = {
                ...json,
                lookup: {
                  ...lu,
                  status: 'resolved',
                  japaneseTitle: jp,
                  provider: lu.provider ? `${lu.provider}+mangaupdatesClient` : 'mangaUpdates(extension)',
                  trace: `${lu.trace || ''} | mangaupdates:direct_hit`.replace(/^\s*\|\s*/, '').trim(),
                },
              };
            }
          } catch (_) {
            /* noop */
          }
        }
      }
    }
    if (json?.lookup && typeof globalThis.applyJapaneseTitleLookupToProduct === 'function') {
      return globalThis.applyJapaneseTitleLookupToProduct({ ...product, titleAnalysis: payload.titleAnalysis }, json);
    }
    return { ...product, titleAnalysis: payload.titleAnalysis, japaneseTitleLookup: json?.lookup || product.japaneseTitleLookup || null };
  } catch (error) {
    return {
      ...product,
      titleAnalysis: payload.titleAnalysis,
      japaneseTitleLookup: {
        status: 'failed',
        provider: 'extension',
        normalizedSearchTitle: payload.titleAnalysis.normalizedSearchTitle || '',
        extractedWorkTitle: payload.titleAnalysis.extractedWorkTitle || '',
        trace: 'extension:lookup_failed',
        errors: [error?.message || String(error || 'lookup failed')],
      }
    };
  }
}

async function postProductToSheet(gasUrl, product, fallbackGenre) {
  const payload = buildSheetPayload(product, fallbackGenre);
  let json = await postAladinPayloadToGas(gasUrl, payload);
  if (shouldFallbackToLegacyAladinAction(json)) {
    json = await postAladinPayloadToGas(gasUrl, buildLegacySheetPayload(product, fallbackGenre));
  }
  if (!json.success && !json.ok) {
    throw new Error(json.error || '書き込み失敗');
  }
  return json;
}

function normalizeProductIsbn(value) {
  return String(value || '').replace(/[^\dXx]/g, '').toUpperCase();
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

  let product = buildProductRecord({
    item,
    itemId,
    pageUrl: tabUrl,
    genre,
    scrapeResult
  });

  product = await lookupJapaneseTitleForProduct(gasUrl, product, genre);
  const mode = await upsertStoredProduct(product);
  await setBackgroundStatus(`画像${product.additionalImages.length}枚取得 / ${genre} → 送信中...`, '');

  const json = await postProductToSheet(gasUrl, product, genre);

  const sheetActionLabel = json.mode === 'created'
    ? '追加'
    : json.mode === 'updated'
      ? '更新'
      : (mode === 'added' ? '追加' : '更新');
  const spreadsheetLabel = json.spreadsheetName || '保存先不明';
  const resolvedGenre = json.resolvedGenre || genre;
  const idLabel = json.spreadsheetId ? ` / ID:${String(json.spreadsheetId).slice(0, 8)}...` : ' / 旧GAS応答の可能性';
  const lookupText = json.lookup && typeof globalThis.renderLookupResult === 'function'
    ? ` / 照会:${globalThis.renderLookupResult(json)}`
    : '';
  await setBackgroundStatus(`✅ ${resolvedGenre} / シート${sheetActionLabel} / ${spreadsheetLabel} / ${json.sheet} 行${json.row}${idLabel}${lookupText}`, 'success');
}

async function processWriteProductsToSheetJob(job) {
  const products = Array.isArray(job.products) ? job.products : [];
  const gasUrl = String(job.gasUrl || '').trim();
  if (!gasUrl) throw new Error('GASのURLを保存してください');
  if (!products.length) throw new Error('シート書込対象の商品がありません');

  let success = 0;
  let failed = 0;
  let skippedDuplicates = 0;
  const errors = [];
  let lastResult = null;
  const seenProductKeys = new Set();

  for (let index = 0; index < products.length; index += 1) {
    let product = products[index] || {};
    const label = product.title || product.itemId || `${index + 1}件目`;
    const itemId = String(product.itemId || '').trim();
    const isbn = normalizeProductIsbn(product.isbn13 || product.isbn);

    if (!itemId) {
      failed += 1;
      errors.push(`${label}: itemId missing`);
      continue;
    }

    const productKeys = [
      itemId ? `item:${itemId}` : '',
      isbn ? `isbn:${isbn}` : ''
    ].filter(Boolean);

    if (productKeys.some(key => seenProductKeys.has(key))) {
      skippedDuplicates += 1;
      continue;
    }
    productKeys.forEach(key => seenProductKeys.add(key));

    await setBackgroundStatus(`⏳ シート書込中 ${index + 1}/${products.length}: ${label}`, '');

    try {
      product = await lookupJapaneseTitleForProduct(gasUrl, product, product.ジャンル || detectGenre(product));
      const result = await postProductToSheet(gasUrl, product, product.ジャンル || detectGenre(product));
      success += 1;
      lastResult = result;
    } catch (error) {
      failed += 1;
      errors.push(`${label}: ${error.message || error}`);
    }
  }

  const type = failed > 0 ? 'error' : 'success';
  const lastSheet = lastResult?.sheet ? ` / 最後:${lastResult.sheet} 行${lastResult.row}` : '';
  const lookupText = lastResult?.lookup && typeof globalThis.renderLookupResult === 'function'
    ? ` / 照会:${globalThis.renderLookupResult(lastResult)}`
    : '';
  const duplicateLabel = skippedDuplicates ? ` / 重複スキップ:${skippedDuplicates}件` : '';
  await setBackgroundStatus(`✅ シート書込 ${success}/${products.length}件${duplicateLabel}${failed ? ` / 失敗:${failed}件` : ''}${lastSheet}${lookupText}`, type);

  return {
    productCount: products.length,
    success,
    failed,
    skippedDuplicates,
    errors,
    lastResult,
  };
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
        } else if (job.type === 'writeProductsToSheet') {
          await processWriteProductsToSheetJob(job);
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
  if (request.action === 'enrichAladinProductTitleLookup') {
    const gasUrl = String(request.gasUrl || '').trim();
    const itemId = String(request.itemId || '').trim();
    lookupJapaneseTitleForProduct(gasUrl, request.product, request.product && request.product.ジャンル)
      .then(async enriched => {
        const stored = await storageGet([STORAGE_KEY]);
        const list = Array.isArray(stored[STORAGE_KEY]) ? [...stored[STORAGE_KEY]] : [];
        const idx = list.findIndex(x => String(x.itemId) === itemId);
        if (idx >= 0) list[idx] = enriched;
        await storageSet({ [STORAGE_KEY]: list });
        sendResponse({ ok: true, product: enriched });
      })
      .catch(error => sendResponse({ ok: false, error: error.message || 'enrich failed' }));
    return true;
  }

  if (request.action === 'prepareAladinImageDownload') {
    prepareAladinImageDownload(request)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ ok: false, error: error.message || 'download prepare failed' }));
    return true;
  }

  if (request.action === 'prepareAladinCsvDownload') {
    prepareAladinCsvDownload(request)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ ok: false, error: error.message || 'csv prepare failed' }));
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

  if (request.action === 'enqueueWriteProductsToSheet') {
    const products = Array.isArray(request.products) ? request.products : [];
    enqueueJob({
      type: 'writeProductsToSheet',
      products,
      gasUrl: request.gasUrl
    }).then(jobId => {
      setBackgroundStatus(`⏳ バックグラウンドで${products.length}件のシート書込を開始しました`, '').catch(() => {});
      sendResponse({ ok: true, jobId, productCount: products.length });
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
