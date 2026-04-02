const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const base = path.join(root, extDir.name);

const manifestFile = path.join(base, 'manifest.json');
const manifest = JSON.parse(fs.readFileSync(manifestFile, 'utf8'));
manifest.host_permissions = manifest.host_permissions || [];
if (!manifest.host_permissions.includes('https://api.mangaupdates.com/*')) {
  manifest.host_permissions.push('https://api.mangaupdates.com/*');
}
fs.writeFileSync(manifestFile, JSON.stringify(manifest, null, 2) + '\n', 'utf8');

const sharedFile = path.join(base, 'popup', 'taiwan', 'popup.shared.js');
let shared = fs.readFileSync(sharedFile, 'utf8');
const sharedMarker = "function escapeRegExp(str) {\n  return String(str || '').replace(/[.*+?^${}()|[\\]\\]/g, '\\\\$&');\n}\n";
if (!shared.includes(sharedMarker)) throw new Error('popup.shared insertion marker not found');
const sharedInsert = [
  "const MANGA_UPDATES_API_BASE = 'https://api.mangaupdates.com/v1';",
  "const mangaUpdatesTitleCache = new Map();",
  "",
  "function normalizeMangaUpdatesTitleKey(value) {",
  "  return extractOriginalTitleText(String(value || ''))",
  "    .toLowerCase()",
  "    .replace(/[\\s\\u3000\"'“”‘’`´・･:：!！?？,，、.．\\-ー‐―–—~〜/／\\\\|()[\\]{}【】「」『』<>]/g, '');",
  "}",
  "",
  "function hasJapaneseTitleSignal(value) {",
  "  return /[ぁ-ゖァ-ヺ々〆ヵヶ]/u.test(String(value || ''));",
  "}",
  "",
  "function uniqNonEmptyTitles(values) {",
  "  return [...new Set((values || []).map(trimValue).filter(Boolean))];",
  "}",
  "",
  "function shouldUseMangaUpdatesLookup(product) {",
  "  if (!product) return false;",
  "  return true;",
  "}",
  "",
  "function buildMangaUpdatesLookupQueries(product) {",
  "  const sheetType = typeof getProductSheetType === 'function' ? getProductSheetType(product) : '';",
  "  const rawTitle = trimValue(product?.商品名 || '');",
  "  const goodsWorkTitle = sheetType === 'goods' && typeof extractGoodsWorkTitleText === 'function'",
  "    ? trimValue(product?.作品名原題 || extractGoodsWorkTitleText(rawTitle))",
  "    : trimValue(product?.作品名原題 || '');",
  "",
  "  return uniqNonEmptyTitles([",
  "    goodsWorkTitle,",
  "    trimValue(product?.原題タイトル || ''),",
  "    extractOriginalTitleText(rawTitle),",
  "  ]).filter(title => normalizeMangaUpdatesTitleKey(title).length >= 2);",
  "}",
  "",
  "async function fetchMangaUpdatesJson(endpoint, init = {}) {",
  "  const response = await fetch(`${MANGA_UPDATES_API_BASE}${endpoint}`, {",
  "    headers: {",
  "      'Accept': 'application/json',",
  "      'Content-Type': 'application/json',",
  "      ...(init.headers || {}),",
  "    },",
  "    ...init,",
  "  });",
  "",
  "  if (!response.ok) {",
  "    throw new Error(`MangaUpdates API error: ${response.status}`);",
  "  }",
  "",
  "  return response.json();",
  "}",
  "",
  "function isStrongMangaUpdatesMatch(query, searchResult, detail) {",
  "  const queryKey = normalizeMangaUpdatesTitleKey(query);",
  "  if (!queryKey) return false;",
  "",
  "  const candidates = uniqNonEmptyTitles([",
  "    searchResult?.hit_title,",
  "    searchResult?.record?.title,",
  "    ...(detail?.associated || []).map(item => item?.title || ''),",
  "  ]);",
  "",
  "  return candidates.some(candidate => normalizeMangaUpdatesTitleKey(candidate) === queryKey);",
  "}",
  "",
  "function pickJapaneseMangaUpdatesTitle(searchResult, detail) {",
  "  const candidates = uniqNonEmptyTitles([",
  "    searchResult?.hit_title,",
  "    ...(detail?.associated || []).map(item => item?.title || ''),",
  "    detail?.title,",
  "    searchResult?.record?.title,",
  "  ]);",
  "",
  "  for (const candidate of candidates) {",
  "    if (!hasJapaneseTitleSignal(candidate)) continue;",
  "    const normalized = extractOriginalTitleText(candidate) || trimValue(candidate);",
  "    if (normalized) return normalized;",
  "  }",
  "",
  "  return '';",
  "}",
  "",
  "async function lookupJapaneseTitleViaMangaUpdates(query) {",
  "  const cacheKey = normalizeMangaUpdatesTitleKey(query);",
  "  if (!cacheKey) return '';",
  "  if (mangaUpdatesTitleCache.has(cacheKey)) return mangaUpdatesTitleCache.get(cacheKey) || '';",
  "",
  "  const searchResponse = await fetchMangaUpdatesJson('/series/search', {",
  "    method: 'POST',",
  "    body: JSON.stringify({ search: query }),",
  "  });",
  "",
  "  const results = Array.isArray(searchResponse?.results) ? searchResponse.results.slice(0, 3) : [];",
  "  for (const result of results) {",
  "    const seriesId = result?.record?.series_id;",
  "    if (!seriesId) continue;",
  "",
  "    const detail = await fetchMangaUpdatesJson(`/series/${seriesId}`);",
  "    if (!isStrongMangaUpdatesMatch(query, result, detail)) continue;",
  "",
  "    const japaneseTitle = pickJapaneseMangaUpdatesTitle(result, detail);",
  "    if (japaneseTitle) {",
  "      mangaUpdatesTitleCache.set(cacheKey, japaneseTitle);",
  "      return japaneseTitle;",
  "    }",
  "  }",
  "",
  "  mangaUpdatesTitleCache.set(cacheKey, '');",
  "  return '';",
  "}",
  "",
  "async function enrichProductWithMangaUpdatesJapaneseTitle(product) {",
  "  if (!shouldUseMangaUpdatesLookup(product)) return product;",
  "",
  "  const sheetType = typeof getProductSheetType === 'function' ? getProductSheetType(product) : '';",
  "  const needsJapaneseTitle = !trimValue(product?.日本語タイトル || '');",
  "  const needsWorkJapaneseTitle = sheetType === 'goods' && !trimValue(product?.作品名日本語 || '');",
  "  if (!needsJapaneseTitle && !needsWorkJapaneseTitle) return product;",
  "",
  "  const queries = buildMangaUpdatesLookupQueries(product);",
  "  for (const query of queries) {",
  "    const japaneseTitle = await lookupJapaneseTitleViaMangaUpdates(query);",
  "    if (!japaneseTitle) continue;",
  "",
  "    if (needsJapaneseTitle && !trimValue(product?.日本語タイトル || '')) {",
  "      product.日本語タイトル = japaneseTitle;",
  "    }",
  "    if (needsWorkJapaneseTitle && !trimValue(product?.作品名日本語 || '')) {",
  "      product.作品名日本語 = japaneseTitle;",
  "    }",
  "    break;",
  "  }",
  "",
  "  return product;",
  "}",
  "",
].join('\n') + sharedMarker;
shared = shared.replace(sharedMarker, sharedInsert);
fs.writeFileSync(sharedFile, shared, 'utf8');

const popupFile = path.join(base, 'popup', 'taiwan', 'popup.js');
let popup = fs.readFileSync(popupFile, 'utf8');
const oldPopup = `async function handleAddCurrentPage() {\n  setStatus('取得中...', '');\n  const data = await requestActiveTabProductInfo();\n  if (products.some(p => extractProductCode(p) === extractProductCode(data))) {\n    throw new Error('この商品はすでにリストにあります');\n  }\n\n  products.push(data);`;
const newPopup = `async function handleAddCurrentPage() {\n  setStatus('取得中...', '');\n  let data = await requestActiveTabProductInfo();\n  if (products.some(p => extractProductCode(p) === extractProductCode(data))) {\n    throw new Error('この商品はすでにリストにあります');\n  }\n\n  try {\n    data = await enrichProductWithMangaUpdatesJapaneseTitle(data);\n  } catch (error) {\n    console.warn('MangaUpdates title enrichment failed:', error);\n  }\n\n  products.push(data);`;
if (!popup.includes(oldPopup)) throw new Error('popup add handler block not found');
popup = popup.replace(oldPopup, newPopup);
fs.writeFileSync(popupFile, popup, 'utf8');

console.log(manifestFile);
console.log(sharedFile);
console.log(popupFile);
