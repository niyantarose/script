// popup.js - 台湾 books.com.tw 拡張 UI / 送信制御

const DEFAULT_STATUS_MESSAGE = 'books.com.tw の商品ページで「追加」を押してください';
const GAS_SYNC_SUCCESS_MESSAGE_TTL_MS = 12000;
const GAS_SYNC_ERROR_MESSAGE_TTL_MS = 8000;
const DEFAULT_SEND_BUTTON_LABEL = 'シートに書込';
const BUSY_SEND_BUTTON_LABEL = '書き込み中...';
let gasSyncSuccessClearTimer = null;
let sendToSheetRequestInFlight = false;

// ===== ログ機能 =====
const POPUP_LOG_MAX = 500;
const POPUP_LOG_STORAGE_KEY = 'popupLogLines_v1';
const popupLogLines = [];
let lastPopupStatusLogMessage = '';

function _logTime(date = new Date()) {
  const hh = String(date.getHours()).padStart(2, '0');
  const mm = String(date.getMinutes()).padStart(2, '0');
  const ss = String(date.getSeconds()).padStart(2, '0');
  const ms = String(date.getMilliseconds()).padStart(3, '0');
  return `${hh}:${mm}:${ss}.${ms}`;
}

function _logHtmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function renderLogPanel() {
  const body = document.getElementById('log-body');
  if (!body) return;
  if (!popupLogLines.length) {
    body.textContent = '（まだログはありません）';
    return;
  }
  body.innerHTML = popupLogLines.map(entry => {
    const cls = entry.type ? `log-${entry.type}` : '';
    return `<div class="${cls}"><span class="log-time">${_logHtmlEscape(entry.time)}</span>${_logHtmlEscape(entry.message)}</div>`;
  }).join('');
  body.scrollTop = body.scrollHeight;
}

function appendPopupLog(message, type = 'info') {
  const text = String(message || '').trim();
  if (!text) return;
  const entry = { time: _logTime(), type, message: text };
  popupLogLines.push(entry);
  if (popupLogLines.length > POPUP_LOG_MAX) {
    popupLogLines.splice(0, popupLogLines.length - POPUP_LOG_MAX);
  }
  renderLogPanel();
  try {
    chrome.storage?.local?.set({ [POPUP_LOG_STORAGE_KEY]: popupLogLines });
  } catch (_) { /* ignore */ }
  const consoleMethod = type === 'error' ? 'error' : type === 'warn' ? 'warn' : 'log';
  console[consoleMethod](`[popup] ${text}`);
}

async function loadStoredLogs() {
  try {
    const result = await new Promise(resolve => {
      chrome.storage.local.get(POPUP_LOG_STORAGE_KEY, r => resolve(r || {}));
    });
    const stored = Array.isArray(result[POPUP_LOG_STORAGE_KEY]) ? result[POPUP_LOG_STORAGE_KEY] : [];
    popupLogLines.length = 0;
    popupLogLines.push(...stored.slice(-POPUP_LOG_MAX));
    renderLogPanel();
  } catch (_) { /* ignore */ }
}

function toggleLogPanel(force) {
  const panel = document.getElementById('log-panel');
  if (!panel) return;
  const open = typeof force === 'boolean' ? force : !panel.classList.contains('open');
  panel.classList.toggle('open', open);
  if (open) renderLogPanel();
}

function clearPopupLog() {
  popupLogLines.length = 0;
  renderLogPanel();
  try {
    chrome.storage?.local?.remove(POPUP_LOG_STORAGE_KEY);
  } catch (_) { /* ignore */ }
}

async function copyPopupLogToClipboard() {
  const text = popupLogLines
    .map(e => `${e.time} [${e.type || 'info'}] ${e.message}`)
    .join('\n');
  try {
    await navigator.clipboard.writeText(text || '(empty)');
    appendPopupLog('ログをクリップボードにコピーしました', 'success');
  } catch (error) {
    appendPopupLog(`ログコピー失敗: ${error?.message || error}`, 'error');
  }
}

window.addEventListener('error', event => {
  appendPopupLog(`window.error: ${event.message} @ ${event.filename}:${event.lineno}:${event.colno}`, 'error');
});
window.addEventListener('unhandledrejection', event => {
  const reason = event.reason;
  const msg = reason?.message || reason?.toString?.() || String(reason);
  appendPopupLog(`unhandledrejection: ${msg}`, 'error');
});

async function ensurePopupSharedReady() {
  const requiredNames = ['storageGet', 'storageSet', 'refreshGasSyncStatus', 'clearGasSyncStatus', 'resetGasSyncState'];
  if (requiredNames.every(name => typeof globalThis[name] === 'function')) {
    return;
  }

  await new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'popup.shared.js';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('popup.shared.js の読み込みに失敗しました'));
    document.head.appendChild(script);
  });

  if (!requiredNames.every(name => typeof globalThis[name] === 'function')) {
    throw new Error('popup.shared.js の初期化に失敗しました');
  }
}

async function initPopup() {
  try {
    await ensurePopupSharedReady();
    const result = await storageGet(['products', MANUAL_SAVE_KEY, GAS_URL_KEY, GAS_URL_DRAFT_KEY]);
    products = result.products || [];
    manualSaveImages = true;
    gasWebAppUrl = String(result[GAS_URL_DRAFT_KEY] || result[GAS_URL_KEY] || '').trim();

    const manualSaveCheckbox = document.getElementById('opt-manual-save');
    if (manualSaveCheckbox) {
      manualSaveCheckbox.checked = manualSaveImages;
      manualSaveCheckbox.addEventListener('change', async () => {
        const nextValue = !!manualSaveCheckbox.checked;
        manualSaveImages = nextValue;
        try {
          await storageSet({ [MANUAL_SAVE_KEY]: manualSaveImages });
        } catch (error) {
          manualSaveImages = !nextValue;
          manualSaveCheckbox.checked = manualSaveImages;
          setStatus(`❌ ${error.message}`, 'error');
        }
      });
    }

    const gasUrlInput = document.getElementById('gas-url');
    if (gasUrlInput) {
      gasUrlInput.value = gasWebAppUrl;
      gasUrlInput.addEventListener('input', () => {
        scheduleGasUrlDraftSave(gasUrlInput.value, { updateUi: true, saveCanonical: false, delayMs: 120 });
      });
      gasUrlInput.addEventListener('change', () => {
        persistGasUrlValue(gasUrlInput.value, { saveCanonical: isValidGasUrl(gasUrlInput.value) }).catch(error => {
          console.error('Failed to persist GAS URL:', error);
        });
      });
      gasUrlInput.addEventListener('blur', () => {
        persistGasUrlValue(gasUrlInput.value, { saveCanonical: isValidGasUrl(gasUrlInput.value) }).catch(error => {
          console.error('Failed to persist GAS URL:', error);
        });
      });
    }

    // GASコールドスタート防止：ポップアップ起動時にウォームアップGETを先行送信
    if (isValidGasUrl(gasWebAppUrl)) {
      warmupGas(gasWebAppUrl);
    }

    renderList();
    await refreshGasSyncStatus({ updateUi: true, applyMessage: true });
    if (gasSyncStatus?.state && !['queued', 'running'].includes(gasSyncStatus.state)) {
      scheduleGasSyncSuccessAutoClear(gasSyncStatus);
    }
    updateUI();
    if (gasSyncStatus?.state && ['queued', 'running'].includes(gasSyncStatus.state)) {
      startGasSyncStatusPolling();
    }
  } catch (error) {
    if (typeof setStatus === 'function') {
      setStatus(`❌ ${error.message}`, 'error');
    } else {
      const status = document.getElementById('status');
      if (status) {
        status.textContent = `❌ ${error.message}`;
      }
      console.error(error);
    }
  }
}

function updateUI() {
  const hasProducts = products.length > 0;
  const gasBusy = isGasSyncBusyStatus(gasSyncStatus);
  const sendBusy = gasBusy || sendToSheetRequestInFlight;
  document.getElementById('item-count').textContent = `${products.length} 件`;
  document.getElementById('btn-download').disabled = !hasProducts;
  document.getElementById('btn-clear').disabled = !hasProducts;

  const imageOnlyButton = document.getElementById('btn-download-images');
  if (imageOnlyButton) {
    imageOnlyButton.disabled = !hasProducts;
  }

  const sendButton = document.getElementById('btn-send-sheet');
  if (sendButton) {
    sendButton.disabled = !hasProducts || !isValidGasUrl(gasWebAppUrl) || sendBusy;
    sendButton.textContent = sendBusy ? BUSY_SEND_BUTTON_LABEL : DEFAULT_SEND_BUTTON_LABEL;
  }
}

function renderList() {
  const container = document.getElementById('list-container');
  if (products.length === 0) {
    container.innerHTML = '<div class="empty-state">まだ商品が追加されていません</div>';
    return;
  }

  container.innerHTML = products.map((p, i) => `
    <div class="product-item">
      ${getProductThumbnailUrl(p) ? `<img class="product-img" src="${escapeHtml(getProductThumbnailUrl(p))}">` : '<div class="product-img"></div>'}
      <div class="product-info">
        <div class="product-name" title="${p.商品名 || ''}">${p.商品名 || '（商品名なし）'}</div>
        <div class="product-meta">
          <span class="badge ${getProductBadgeClass(p)}">${getProductKindLabel(p)}</span>
          <span class="price">NT$ ${p.価格 || '-'}</span>
          <span>${extractProductCode(p) || ''}</span>
        </div>
        ${renderTitleAnalysisPreview(p)}
      </div>
      <button class="btn-remove" data-index="${i}" title="削除">×</button>
    </div>
  `).join('');

  container.querySelectorAll('.btn-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.index, 10);
      products.splice(idx, 1);
      save();
      renderList();
      updateUI();
    });
  });

  container.querySelectorAll('.product-img').forEach(img => {
    if (img.tagName !== 'IMG') return;
    img.addEventListener('error', () => {
      img.style.display = 'none';
    }, { once: true });
  });
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderTitleAnalysisPreview(product) {
  if (typeof analyzeProductTitle !== 'function') return '';
  const analysis = product?.titleAnalysis || analyzeProductTitle({ ...product, source: 'books_tw' });
  const lookupText = product?.japaneseTitleLookup && typeof renderLookupResult === 'function'
    ? `照会: ${renderLookupResult({ lookup: product.japaneseTitleLookup })}`
    : '';
  const parts = [
    analysis.itemType,
    analysis.normalizedSearchTitle ? `検索: ${analysis.normalizedSearchTitle}` : '',
    analysis.extractedWorkTitle ? `作品: ${analysis.extractedWorkTitle}` : '',
    analysis.volume ? `巻: ${analysis.volume}` : '',
    analysis.edition ? `版: ${analysis.edition}` : '',
    analysis.goodsType ? `グッズ: ${analysis.goodsType}` : '',
    analysis.magazineName ? `雑誌: ${analysis.magazineName}${analysis.year ? ` ${analysis.year}` : ''}${analysis.month ? `.${analysis.month}` : ''}` : '',
    lookupText,
    (analysis.warnings || []).length ? `警告: ${analysis.warnings.join(' / ')}` : '',
  ].filter(Boolean).map(escapeHtml);
  return parts.length ? `<div class="product-analysis" title="${parts.join(' / ')}">${parts.join(' / ')}</div>` : '';
}

function warmupGas(url) {
  // GASへGETリクエストを送ってコールドスタートを解消しておく（fire-and-forget）
  const controller = new AbortController();
  setTimeout(() => controller.abort(), 15000);
  fetch(url, { method: 'GET', credentials: 'omit', redirect: 'follow', signal: controller.signal })
    .catch(() => {});
}

function setStatus(msg, type = '', options = {}) {
  const el = document.getElementById('status');
  const text = msg || DEFAULT_STATUS_MESSAGE;
  el.textContent = text;
  el.className = `status-bar ${type}`;
  if (!text || text === DEFAULT_STATUS_MESSAGE) {
    lastPopupStatusLogMessage = '';
    return;
  }
  const logType = type === 'error' ? 'error' : type === 'success' ? 'success' : 'info';
  const forceLog = options.forceLog === true;
  if (forceLog || text !== lastPopupStatusLogMessage) {
    appendPopupLog(text, logType);
    lastPopupStatusLogMessage = text;
  }
}

function clearGasSyncSuccessAutoClearTimer() {
  if (!gasSyncSuccessClearTimer) return;
  clearTimeout(gasSyncSuccessClearTimer);
  gasSyncSuccessClearTimer = null;
}

function restoreDefaultStatusIfMatching(previousStatus) {
  const previousMessage = gasSyncStatusMessage(previousStatus);
  const statusEl = document.getElementById('status');
  if (!previousMessage || !statusEl || statusEl.textContent !== previousMessage) return;
  setStatus(DEFAULT_STATUS_MESSAGE, '');
}

function syncGasStatusUi(previousStatus, nextStatus) {
  const previousMessage = gasSyncStatusMessage(previousStatus);
  gasSyncStatus = nextStatus || null;
  const nextMessage = gasSyncStatusMessage(gasSyncStatus);

  if (nextMessage && nextMessage !== previousMessage) {
    setStatus(nextMessage, gasSyncStatusType(gasSyncStatus));
  } else if (!nextMessage && previousStatus) {
    restoreDefaultStatusIfMatching(previousStatus);
  }

  updateUI();
  if (isGasSyncBusyStatus(gasSyncStatus)) {
    startGasSyncStatusPolling();
  } else {
    stopGasSyncStatusPolling();
  }
}

function scheduleGasSyncSuccessAutoClear(status) {
  clearGasSyncSuccessAutoClearTimer();
  if (status?.state !== 'success' && status?.state !== 'error') return;

  const ttl = status.state === 'error' ? GAS_SYNC_ERROR_MESSAGE_TTL_MS : GAS_SYNC_SUCCESS_MESSAGE_TTL_MS;
  const expectedState = status.state;

  gasSyncSuccessClearTimer = window.setTimeout(async () => {
    if (gasSyncStatus?.jobId !== status.jobId || gasSyncStatus?.state !== expectedState) return;

    try {
      await clearGasSyncStatus();
    } catch (error) {
      console.warn('Failed to auto clear GAS sync status:', error);
      const previousStatus = gasSyncStatus;
      syncGasStatusUi(previousStatus, null);
    }
    updateUI();
  }, ttl);
}

function handleGasSyncStorageChange(changes, areaName) {
  if (areaName !== 'local') return;
  if (!Object.prototype.hasOwnProperty.call(changes, GAS_SYNC_STATUS_KEY)) return;

  const previousStatus = changes[GAS_SYNC_STATUS_KEY]?.oldValue || gasSyncStatus || null;
  const nextStatus = changes[GAS_SYNC_STATUS_KEY]?.newValue || null;

  clearGasSyncSuccessAutoClearTimer();
  syncGasStatusUi(previousStatus, nextStatus);
  scheduleGasSyncSuccessAutoClear(nextStatus);
}

chrome.storage.onChanged.addListener(handleGasSyncStorageChange);

function save() {
  storageSet({ products }).catch(error => {
    console.error('Failed to save products:', error);
  });
}

function productForSheetRowBuild_(product) {
  // 商品ページのタイトル直下から取得できた日本語タイトルが最優先（最も信頼できる）。
  const pageTrustedJp = String(product?.日本語タイトル取得元 || '') === 'page_trusted'
    ? String(product?.日本語タイトル || '').trim()
    : '';
  if (pageTrustedJp) {
    return {
      ...product,
      日本語タイトル: pageTrustedJp,
      作品名日本語: product.作品名日本語 || pageTrustedJp,
      '作品名（日本語）': product['作品名（日本語）'] || pageTrustedJp,
    };
  }
  const lookup = product?.japaneseTitleLookup || null;
  if (lookup?.status === 'resolved' && String(lookup.japaneseTitle || '').trim()) {
    const jp = String(lookup.japaneseTitle).trim();
    return {
      ...product,
      日本語タイトル: jp,
      作品名日本語: product.作品名日本語 || jp,
      '作品名（日本語）': product['作品名（日本語）'] || jp,
    };
  }
  const analysis = product?.titleAnalysis || (typeof analyzeProductTitle === 'function'
    ? analyzeProductTitle({ ...product, source: 'books_tw' })
    : null);
  return stripInvalidPageJapaneseTitle_(product, analysis);
}

function buildGasPayload(items) {
  return {
    action: 'upsertProductWithLookup',
    source: 'books_tw',
    appendMode: 'append',
    requestedAt: new Date().toISOString(),
    items: items.map(product => {
      const kind = getProductSheetType(product);
      const titleAnalysis = product.titleAnalysis || (typeof analyzeProductTitle === 'function'
        ? analyzeProductTitle({ ...product, source: 'books_tw' })
        : null);
      const lookup = product.japaneseTitleLookup || null;
      const lookupFailedInExtension = lookup
        && lookup.status === 'failed'
        && String(lookup.provider || '').trim() === 'extension';
      const rowProduct = productForSheetRowBuild_(product);
      return {
        source: 'books_tw',
        rawItem: product,
        productCode: extractProductCode(product),
        itemType: kind,
        sheetName: getProductSheetName(product),
        headers: kind === 'magazine' ? MAGAZINE_SHEET_HEADERS : kind === 'book' ? BOOK_SHEET_HEADERS : GOODS_SHEET_HEADERS,
        rowData: kind === 'magazine' ? buildMagazineSheetRow(rowProduct) : kind === 'book' ? buildBookSheetRow(rowProduct) : buildGoodsSheetRow(rowProduct),
        titleAnalysis,
        japaneseTitleLookup: lookupFailedInExtension ? null : lookup,
      };
    }),
  };
}

function stripInvalidPageJapaneseTitle_(product, analysis) {
  if (!product || typeof product !== 'object' || !analysis) return product;
  const jp = String(product.日本語タイトル || '').trim();
  if (!jp) return product;
  // 商品ページのタイトル直下から取得できた日本語タイトルは最も信頼できるので検証せず採用する。
  if (String(product.日本語タイトル取得元 || '') === 'page_trusted') return product;
  if (typeof validateJapaneseTitleAgainstQuery !== 'function'
    || typeof buildMuQueryVariants !== 'function') {
    return product;
  }
  const queries = buildMuQueryVariants(analysis, product);
  if (validateJapaneseTitleAgainstQuery(jp, queries, 'page_subtitle')) return product;
  const next = { ...product, 日本語タイトル: '' };
  for (const key of ['作品名日本語', '作品名（日本語）', '商品名日本語', '商品名（日本語）']) {
    if (String(next[key] || '').trim() === jp) next[key] = '';
  }
  return next;
}

function clearPendingJapaneseTitleMarkers(product) {
  if (!product || typeof product !== 'object') return product;
  const cleaned = { ...product };
  for (const key of ['日本語タイトル', '作品名日本語', '作品名（日本語）', '商品名日本語', '商品名（日本語）']) {
    const value = String(cleaned[key] || '').trim();
    if (!value) continue;
    if ((typeof isMangaUpdatesStatusValue === 'function' && isMangaUpdatesStatusValue(value))
      || /^(?:登録なし|照会失敗|未照会)(?:\s*[（(]|$)/u.test(value)) {
      cleaned[key] = '';
    }
  }
  if (typeof analyzeProductTitle === 'function') {
    cleaned.titleAnalysis = analyzeProductTitle({ ...cleaned, source: 'books_tw' });
  }
  return cleaned;
}

function hasReusableJapaneseTitleLookup(product, analysis) {
  const lookup = product?.japaneseTitleLookup;
  if (!lookup || typeof lookup !== 'object') return false;
  const lookupKey = String(lookup.normalizedSearchTitle || '').trim();
  const analysisKey = String(analysis?.normalizedSearchTitle || '').trim();
  return Boolean(lookup.status && lookupKey && analysisKey && lookupKey === analysisKey);
}

async function enrichProductWithJapaneseTitleLookup(product, options = {}) {
  if (!product || typeof product !== 'object') return product;
  if (typeof analyzeProductTitle !== 'function') return product;

  const analysis = product.titleAnalysis || analyzeProductTitle({ ...product, source: 'books_tw' });
  let nextProduct = { ...product, titleAnalysis: analysis };
  if (!options.force && hasReusableJapaneseTitleLookup(nextProduct, analysis)) return nextProduct;
  if (typeof lookupJapaneseTitle !== 'function') return nextProduct;

  const effectiveUrl = trimValue(options.gasUrl || gasWebAppUrl || '');
  if (!isValidGasUrl(effectiveUrl)) return nextProduct;

  try {
    const result = await lookupJapaneseTitle(analysis, {
      url: effectiveUrl,
      gasUrl: effectiveUrl,
      rawItem: nextProduct,
    });
    nextProduct = typeof applyJapaneseTitleLookupToProduct === 'function'
      ? applyJapaneseTitleLookupToProduct(nextProduct, result)
      : { ...nextProduct, japaneseTitleLookup: result.lookup || null };
  } catch (error) {
    nextProduct.japaneseTitleLookup = {
      status: 'failed',
      provider: 'extension',
      normalizedSearchTitle: analysis.normalizedSearchTitle || '',
      extractedWorkTitle: analysis.extractedWorkTitle || '',
      trace: 'extension:lookup_failed',
      errors: [error?.message || String(error || 'lookup failed')],
    };
  }

  return nextProduct;
}

async function prepareProductsForSheetSend(items, options = {}) {
  const sourceItems = Array.isArray(items) ? items : [];
  const prepared = [];
  for (const item of sourceItems) {
    let product = clearPendingJapaneseTitleMarkers(item);
    const analysis = product.titleAnalysis || (typeof analyzeProductTitle === 'function'
      ? analyzeProductTitle({ ...product, source: 'books_tw' })
      : null);
    product = stripInvalidPageJapaneseTitle_(product, analysis);
    const lookup = product.japaneseTitleLookup;
    const lookupJp = String(lookup?.japaneseTitle || '').trim();
    const needsEnrich = !lookup
      || (lookup.status === 'resolved' && lookupJp
        && typeof validateJapaneseTitleAgainstQuery === 'function'
        && typeof buildMuQueryVariants === 'function'
        && !validateJapaneseTitleAgainstQuery(lookupJp, buildMuQueryVariants(analysis, product), lookup.provider || ''))
      || (!lookupJp && lookup?.status !== 'series_found_no_japanese' && lookup?.status !== 'not_found');
    if (needsEnrich && typeof enrichProductWithJapaneseTitleLookup === 'function') {
      product = await enrichProductWithJapaneseTitleLookup(product, {
        gasUrl: options.gasUrl || gasWebAppUrl,
        force: true,
      });
    }
    prepared.push(product);
  }
  return prepared;
}

function isValidGasUrl(url) {
  if (typeof isGasWebAppExecUrl === 'function') {
    return isGasWebAppExecUrl(url);
  }
  try {
    const parsed = new URL(String(url || '').trim());
    return parsed.protocol === 'https:'
      && parsed.hostname === 'script.google.com'
      && /^\/macros\/s\/[^/]+\/exec\/?$/i.test(parsed.pathname);
  } catch {
    return false;
  }
}

function buildCsvBaseRow(product) {
  return [
    getProductKindLabel(product),
    extractProductCode(product),
    product.商品名 || '',
    product.価格 || '',
    product.発売日 || '',
    product.画像URL || '',
    getAdditionalImagesValue(product),
    product.URL || '',
    product.商品説明 || '',
  ];
}

async function handleAddCurrentPage() {
  setStatus('取得中...', '');
  let data = await requestActiveTabProductInfo();
  const existingIndex = products.findIndex(p => extractProductCode(p) === extractProductCode(data));
  if (existingIndex >= 0) {
    products[existingIndex] = { ...products[existingIndex], ...data };
    save();
    renderList();
    updateUI();
    setStatus('✅ 既存商品の情報を更新しました: ' + (data.商品名?.slice(0, 30) || extractProductCode(data)), 'success');
    return;
  }

  data = clearPendingJapaneseTitleMarkers(data);

  products.push(data);
  save();
  renderList();
  await refreshGasSyncStatus({ updateUi: false, applyMessage: false });
  updateUI();
  if (gasSyncStatus?.state && ['queued', 'running'].includes(gasSyncStatus.state)) {
    startGasSyncStatusPolling();
  }
  const code = extractProductCode(data);
  // 追加処理の体感を優先し、MU照会は後段で非同期実行する
  (async () => {
    try {
      const idx = products.findIndex(item => extractProductCode(item) === code);
      if (idx < 0) return;
      let enriched = { ...products[idx] };
      if (typeof enrichProductWithJapaneseTitleLookup === 'function') {
        enriched = await enrichProductWithJapaneseTitleLookup(enriched, { gasUrl: gasWebAppUrl });
      } else if (typeof enrichProductWithMangaUpdatesJapaneseTitle === 'function') {
        enriched = await enrichProductWithMangaUpdatesJapaneseTitle(enriched);
      }
      products[idx] = { ...products[idx], ...enriched };
      save();
      renderList();
    } catch (error) {
      console.warn('[popup] MU enrichment エラー:', error);
    }
  })();
  setStatus(
    `✅ ${getProductKindLabel(data)}を追加（照会はバックグラウンド）: ${data.商品名?.slice(0, 30) || code}`,
    'success'
  );
}

async function handleDownloadImagesOnly() {
  if (products.length === 0) return;
  setStatus('保存準備中...','');

  const directSaveContext = null;
  setStatus('画像を保存中...','');

  const imageResult = await downloadProductImages(products, {
    manualSave: true,
    directSaveContext,
  });
  const destinationLabel = describeSaveDestination(
    directSaveContext,
    '手動保存ダイアログ'
  );
  const firstFailedDetail = Array.isArray(imageResult.failedDetails) ? imageResult.failedDetails[0] : null;
  const failedReasonText = firstFailedDetail
    ? ` / 例: ${firstFailedDetail.productCode || '-'} ${firstFailedDetail.error || ''}`
    : '';
  setStatus(
    `✅ 画像:${imageResult.success}/${imageResult.total}件保存 / フォルダ:${imageResult.folders.length} / ${destinationLabel}` +
      (imageResult.failed > 0 ? `（失敗:${imageResult.failed}${failedReasonText}）` : ''),
    imageResult.failed > 0 ? 'error' : 'success'
  );
}

async function handleDownloadCsvAndImages() {
  if (products.length === 0) return;
  setStatus('保存準備中...','');

  const directSaveContext = null;
  setStatus('CSVと画像を出力中...','');

  const baseHeaders = [
    '種別',
    '商品コード',
    '商品名',
    '価格',
    '発売日',
    '画像URL',
    '追加画像URL',
    '商品ページURL',
    '商品説明',
  ];
  const bookHeaders = ['ISBN', '著者', '翻訳者', 'イラストレーター', '出版社'];
  const magazineHeaders = ['原題タイトル', '原題商品名'];

  const books = products.filter(product => getProductSheetType(product) === 'book');
  const magazines = products.filter(product => getProductSheetType(product) === 'magazine');
  const goods = products.filter(product => getProductSheetType(product) === 'goods');

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  let exported = 0;

  if (books.length > 0) {
    const rows = books.map(buildBookCsvRow);
    const ok = await downloadCsvFile(
      [...baseHeaders, ...bookHeaders],
      rows,
      buildCsvDownloadPath(`bokukuro_${dateStr}_books.csv`),
      { directSaveContext }
    );
    if (ok) exported += 1;
  }

  if (magazines.length > 0) {
    const rows = magazines.map(buildMagazineCsvRow);
    const ok = await downloadCsvFile(
      [...baseHeaders, ...magazineHeaders],
      rows,
      buildCsvDownloadPath(`bokukuro_${dateStr}_magazines.csv`),
      { directSaveContext }
    );
    if (ok) exported += 1;
  }

  if (goods.length > 0) {
    const rows = goods.map(buildCsvBaseRow);
    const ok = await downloadCsvFile(
      baseHeaders,
      rows,
      buildCsvDownloadPath(`bokukuro_${dateStr}_goods.csv`),
      { directSaveContext }
    );
    if (ok) exported += 1;
  }

  const imageResult = await downloadProductImages(products, {
    manualSave: true,
    directSaveContext,
  });
  const destinationLabel = describeSaveDestination(
    directSaveContext,
    '手動保存ダイアログ'
  );
  const firstFailedDetail = Array.isArray(imageResult.failedDetails) ? imageResult.failedDetails[0] : null;
  const failedReasonText = firstFailedDetail
    ? ` / 例: ${firstFailedDetail.productCode || '-'} ${firstFailedDetail.error || ''}`
    : '';
  setStatus(
    `✅ CSV:${exported}ファイル / 画像:${imageResult.success}/${imageResult.total}件保存 / フォルダ:${imageResult.folders.length} / ${destinationLabel}` +
      (imageResult.failed > 0 ? `（失敗:${imageResult.failed}${failedReasonText}）` : ''),
    imageResult.failed > 0 ? 'error' : 'success'
  );
}

async function handleSaveGasUrl() {
  const input = document.getElementById('gas-url');
  const url = String(input?.value || '').trim();
  if (!isValidGasUrl(url)) {
    throw new Error('GAS Webアプリの /exec URL を入力してください');
  }

  gasWebAppUrl = url;
  await persistGasUrlValue(gasWebAppUrl, { saveCanonical: true });
  try {
    await clearGasSyncStatus();
  } catch (error) {
    console.warn('Failed to clear GAS sync status after URL save:', error);
  }
  if (input) {
    input.value = gasWebAppUrl;
  }
  updateUI();
  setStatus('✅ GAS URL を保存しました', 'success');
}

async function handleResetGas() {
  stopGasSyncStatusPolling();
  clearGasSyncSuccessAutoClearTimer();
  await resetGasSyncState();
  gasSyncStatus = null;
  updateUI();
  setStatus('✅ GAS送信状態をリセットしました', 'success');
}

async function handleSendToSheet() {
  if (products.length === 0) return;
  if (sendToSheetRequestInFlight || isGasSyncBusyStatus(gasSyncStatus)) return;
  if (!isValidGasUrl(gasWebAppUrl)) {
    throw new Error('先にGAS Webアプリ URL を保存してください');
  }

  sendToSheetRequestInFlight = true;
  updateUI();
  try {
    setStatus('バックグラウンド送信を登録中...', '');
    const response = await sendRuntimeMessage({
      action: 'queueGasPost',
      url: gasWebAppUrl,
      items: buildGasPayload(products).items,
    });

    if (!response?.ok) {
      throw new Error(response?.error || 'バックグラウンド送信の登録に失敗しました');
    }

    clearGasSyncSuccessAutoClearTimer();
    syncGasStatusUi(null, response.status || null);
  } finally {
    sendToSheetRequestInFlight = false;
    updateUI();
  }
}

document.getElementById('btn-save-gas-url').addEventListener('click', async () => {
  try {
    await handleSaveGasUrl();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-reset-gas').addEventListener('click', async () => {
  try {
    await handleResetGas();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-add').addEventListener('click', async () => {
  try {
    await handleAddCurrentPage();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-send-sheet').addEventListener('click', async () => {
  try {
    await handleSendToSheet();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-download').addEventListener('click', async () => {
  try {
    await handleDownloadCsvAndImages();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-download-images').addEventListener('click', async () => {
  try {
    await handleDownloadImagesOnly();
  } catch (error) {
    setStatus(`❌ ${error.message}`, 'error');
  }
});

document.getElementById('btn-clear').addEventListener('click', () => {
  if (products.length === 0) return;
  if (!confirm(`${products.length}件のリストをすべて削除しますか？`)) return;
  products = [];
  save();
  renderList();
  updateUI();
  setStatus('リストをクリアしました', '');
});

document.getElementById('btn-log')?.addEventListener('click', () => {
  toggleLogPanel();
});
document.getElementById('btn-log-close')?.addEventListener('click', () => {
  toggleLogPanel(false);
});
document.getElementById('btn-log-clear')?.addEventListener('click', () => {
  clearPopupLog();
  appendPopupLog('ログをクリアしました', 'info');
});
document.getElementById('btn-log-copy')?.addEventListener('click', () => {
  copyPopupLogToClipboard();
});

loadStoredLogs();
appendPopupLog(`popup起動 v0.2.1 (${navigator.userAgent.replace(/Mozilla\/5\.0\s*/, '').slice(0, 80)})`, 'info');
initPopup();









