// popup.js - 韓国グッズ商品取得ツール

let products = [];
const LAST_STATUS_KEY = 'kr_goods_last_status';

async function initPopup() {
  try {
    const result = await chrome.storage.local.get(['kr_goods_products', LAST_STATUS_KEY]);
    products = result.kr_goods_products || [];
    renderList();
    updateUI();
    const lastStatus = result[LAST_STATUS_KEY];
    if (lastStatus?.message) {
      setStatus(lastStatus.message, lastStatus.type || '');
    }
  } catch (e) {
    setStatus('❌ 初期化失敗: ' + (e?.message || String(e)), 'error');
  }
}

initPopup();

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== 'local') return;
  const statusChange = changes[LAST_STATUS_KEY];
  if (!statusChange || !statusChange.newValue) return;
  const value = statusChange.newValue;
  if (value?.message) {
    setStatus(value.message, value.type || '');
  }
});

function updateUI() {
  document.getElementById('item-count').textContent = `${products.length} 件`;
  document.getElementById('btn-download').disabled = products.length === 0;
  document.getElementById('btn-clear').disabled = products.length === 0;
}

function renderList() {
  const container = document.getElementById('list-container');
  if (products.length === 0) {
    container.innerHTML = '<div class="empty-state">まだ商品が追加されていません</div>';
    return;
  }
  container.innerHTML = products.map((p, i) => `
    <div class="product-item">
      ${p.メイン画像URL ? `<img class="product-img" src="${p.メイン画像URL}" onerror="this.style.display='none'">` : '<div class="product-img"></div>'}
      <div class="product-info">
        <div class="product-name" title="${p.商品名}">${p.商品名 || '（商品名なし）'}</div>
        <div class="product-meta">
          <span class="badge">グッズ</span>
          <span class="shop-code">${p.ショップコード}-${p.商品ID}</span>
          <span class="price">₩${p.価格 || '-'}</span>
        </div>
      </div>
      <button class="btn-remove" data-index="${i}" title="削除">×</button>
    </div>
  `).join('');

  container.querySelectorAll('.btn-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.index);
      products.splice(idx, 1);
      save();
      renderList();
      updateUI();
    });
  });
}

function setStatus(msg, type = '') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = `status-bar ${type}`;
}

function save() {
  chrome.storage.local.set({ kr_goods_products: products });
}

function getDateTimeStr() {
  const now = new Date();
  const d = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  const t = `${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
  return { d, t };
}

function getImageExt(url) {
  const ext = (url.split('.').pop() || '').split('?')[0].toLowerCase();
  if (ext === 'jpeg') return 'jpg';
  return ['jpg', 'png', 'webp', 'gif'].includes(ext) ? ext : 'jpg';
}

function isImageUrl(url) {
  const text = String(url || '').trim();
  if (!/^https?:\/\//i.test(text)) return false;
  if (/image_zoom\d?\.html|\/product\/detail\.html/i.test(text)) return false;
  return /\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?|$)/i.test(text);
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

function extractAdditionalImagesFromHtml(html, pageUrl, mainImageUrl) {
  const urls = [];
  const seen = new Set();
  const mainKey = getDedupeKey(mainImageUrl);

  try {
    const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
    const detailRoot =
      doc.querySelector('.xans-product-additional.left_area') ||
      doc.querySelector('#prdDetail') ||
      doc.querySelector('.xans-product-additional') ||
      doc;

    const images = [
      ...Array.from(detailRoot.querySelectorAll('img[ec-data-src]')),
      ...Array.from(detailRoot.querySelectorAll('img[data-ec-data-src]')),
      ...Array.from(detailRoot.querySelectorAll('.edb-img-tag-w img')),
      ...Array.from(detailRoot.querySelectorAll('img')),
    ];
    const seenElements = new Set();

    images.forEach(img => {
      if (!img || seenElements.has(img)) return;
      seenElements.add(img);

      const styleText = String(img.getAttribute('style') || '').toLowerCase();
      const minHeightMatch = styleText.match(/min-height\s*:\s*(\d+(?:\.\d+)?)px/);
      if (minHeightMatch) {
        const minHeight = parseFloat(minHeightMatch[1]);
        if (!Number.isNaN(minHeight) && minHeight > 0 && minHeight < 180) return;
      }

      const candidates = [
        img.getAttribute('ec-data-src'),
        img.getAttribute('data-ec-data-src'),
        img.getAttribute('data-src'),
        img.getAttribute('data-lazy-src'),
        img.getAttribute('data-original'),
        img.getAttribute('src'),
      ];

      for (const raw of candidates) {
        const abs = toAbsoluteUrl(raw, pageUrl);
        if (!isImageUrl(abs)) continue;
        if (/\/(category|icon|common|layout|banner)\//i.test(abs)) continue;
        const key = getDedupeKey(abs);
        if (!key || key === mainKey || seen.has(key)) continue;
        seen.add(key);
        urls.push(abs);
        break;
      }
    });
  } catch (_e) {
    return [];
  }

  return urls;
}

async function resolveAdditionalImageUrls(product) {
  const fromData = String(product.追加画像URL || '')
    .split(';')
    .map(x => x.trim())
    .filter(isImageUrl);
  if (fromData.length > 0) return fromData;

  const pageUrl = String(product.購入URL || '').trim();
  if (!/^https?:\/\//i.test(pageUrl)) return [];

  try {
    const resp = await fetch(pageUrl, { method: 'GET', credentials: 'include' });
    if (!resp.ok) return [];
    const html = await resp.text();
    return extractAdditionalImagesFromHtml(html, pageUrl, product.メイン画像URL || '');
  } catch (_e) {
    return [];
  }
}

function sanitizePathSegment(segment) {
  return String(segment || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
}

function buildShopProductCode(product) {
  const code = String(product?.ショップコード || product?.ショップ || 'SHOP').trim();
  const id = String(product?.商品ID || '').trim();
  return id ? `${code}-${id}` : code;
}

function buildExportBaseName(items, datePart, timePart) {
  const first = items?.[0] || {};
  const shopName = sanitizePathSegment(first.ショップ || first.ショップコード || 'SHOP');
  const shopProductCode = sanitizePathSegment(buildShopProductCode(first));
  const fallback = `SHOP_${datePart}_${timePart}`;

  if (shopName && shopProductCode) return `${shopName}_${shopProductCode}`;
  if (shopName) return shopName;
  if (shopProductCode) return shopProductCode;
  return fallback;
}

function getCsvFilename(items, datePart, timePart) {
  const base = buildExportBaseName(items, datePart, timePart);
  return `${base}.csv`;
}

// 追加ボタン
document.getElementById('btn-add').addEventListener('click', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  // 対応ショップの商品ページか確認
  let isSupportedPage = false;
  try {
    const u = new URL(String(tab.url || ''));
    const host = u.hostname.toLowerCase();
    const path = u.pathname || '/';
    const isCafe24 =
      /(^|\.)j-meeshop\.com$/.test(host) ||
      /(^|\.)mostore\.co\.kr$/.test(host) ||
      /(^|\.)toonique\.co\.kr$/.test(host) ||
      /(^|\.)webtoonshop\.com$/.test(host);
    if (isCafe24) {
      isSupportedPage = /\/product\//i.test(path);
    } else if (/(^|\.)webtoonfriends\.marpple\.shop$/.test(host)) {
      isSupportedPage = /\/products?\//i.test(path);
    } else if (/(^|\.)smartstore\.naver\.com$/.test(host)) {
      isSupportedPage = /^\/official_w\/products\/.+/i.test(path);
    }
  } catch (_e) {
    isSupportedPage = false;
  }

  if (!isSupportedPage) {
    setStatus('⚠ 対応ショップの商品ページを開いてください（JMEE/MOFUN/TOON/WTSHOP/WTF/OFW）', 'error');
    return;
  }

  setStatus('取得中...', '');

  const requestProductInfo = () => new Promise(resolve => {
    chrome.tabs.sendMessage(tab.id, { action: 'getProductInfo' }, response => {
      if (chrome.runtime.lastError) {
        resolve({ __sendError: chrome.runtime.lastError.message });
        return;
      }
      resolve(response || null);
    });
  });

  try {
    let response = await requestProductInfo();

    // 受信先が無いときだけcontent.jsを1回注入して再試行
    if (!response || response.__sendError) {
      await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['content.js'] });
      response = await requestProductInfo();
      if (response && response.__sendError) {
        throw new Error(response.__sendError);
      }
    }

    handleResponse(response);
  } catch (e) {
    setStatus('❌ 取得失敗: ' + (e?.message || String(e)), 'error');
  }
});

function handleResponse(response) {
  if (!response || !response.success) {
    setStatus('❌ 取得失敗: ' + (response?.error || '不明なエラー'), 'error');
    return;
  }
  const data = response.data;
  if (products.some(p => p.重複チェックキー === data.重複チェックキー)) {
    setStatus('⚠ この商品はすでにリストにあります', 'error');
    return;
  }
  products.push(data);
  save();
  renderList();
  updateUI();
  setStatus(`✅ 追加: ${data.商品名?.slice(0, 30) || data.商品ID}`, 'success');
}

// CSVダウンロード
document.getElementById('btn-download').addEventListener('click', async () => {
  if (products.length === 0) return;
  setStatus('出力ジョブを開始しています...', '');
  const { d, t } = getDateTimeStr();
  try {
    await chrome.runtime.sendMessage({ action: 'startExport', products, datePart: d, timePart: t });
  } catch (e) {
    setStatus(`❌ 出力開始失敗: ${e?.message || String(e)}`, 'error');
    return;
  }
  setStatus('⏳ バックグラウンドで出力中です。閉じても続行します。', '');
});

// クリアボタン
document.getElementById('btn-clear').addEventListener('click', () => {
  if (products.length === 0) return;
  if (!confirm(`${products.length}件のリストをすべて削除しますか？`)) return;
  products = [];
  save();
  renderList();
  updateUI();
  setStatus('リストをクリアしました', '');
});


