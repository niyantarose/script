// popup.js - 台湾 books.com.tw 拡張 UI / 送信制御

const DEFAULT_STATUS_MESSAGE = 'books.com.tw の商品ページで「追加」を押してください';
const GAS_SYNC_SUCCESS_MESSAGE_TTL_MS = 300;
const GAS_SYNC_ERROR_MESSAGE_TTL_MS = 8000;
const DEFAULT_SEND_BUTTON_LABEL = 'シートに書込';
const BUSY_SEND_BUTTON_LABEL = '書き込み中...';
let gasSyncSuccessClearTimer = null;
let sendToSheetRequestInFlight = false;

async function ensurePopupSharedReady() {
  const requiredNames = ['storageGet', 'storageSet', 'refreshGasSyncStatus', 'setStatus'];
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
    await refreshGasSyncStatus({ updateUi: false, applyMessage: false });
    if (gasSyncStatus?.state && !['queued', 'running'].includes(gasSyncStatus.state)) {
      try {
        await clearGasSyncStatus();
      } catch (error) {
        console.warn('Failed to clear stale GAS sync status:', error);
      }
      gasSyncStatus = null;
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
      ${p.画像URL ? `<img class="product-img" src="${p.画像URL}">` : '<div class="product-img"></div>'}
      <div class="product-info">
        <div class="product-name" title="${p.商品名 || ''}">${p.商品名 || '（商品名なし）'}</div>
        <div class="product-meta">
          <span class="badge ${getProductBadgeClass(p)}">${getProductKindLabel(p)}</span>
          <span class="price">NT$ ${p.価格 || '-'}</span>
          <span>${extractProductCode(p) || ''}</span>
        </div>
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

function warmupGas(url) {
  // GASへGETリクエストを送ってコールドスタートを解消しておく（fire-and-forget）
  const controller = new AbortController();
  setTimeout(() => controller.abort(), 15000);
  fetch(url, { method: 'GET', credentials: 'omit', redirect: 'follow', signal: controller.signal })
    .catch(() => {});
}

function setStatus(msg, type = '') {
  const el = document.getElementById('status');
  el.textContent = msg || DEFAULT_STATUS_MESSAGE;
  el.className = `status-bar ${type}`;
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
  gasSyncStatus = nextStatus || null;
  const nextMessage = gasSyncStatusMessage(gasSyncStatus);

  if (nextMessage) {
    setStatus(nextMessage, gasSyncStatusType(gasSyncStatus));
  } else if (previousStatus) {
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

function buildGasPayload(items) {
  return {
    source: 'books.com.tw-extension',
    appendMode: 'append',
    requestedAt: new Date().toISOString(),
    items: items.map(product => {
      const kind = getProductSheetType(product);
      return {
        productCode: extractProductCode(product),
        itemType: kind,
        sheetName: getProductSheetName(product),
        headers: kind === 'magazine' ? MAGAZINE_SHEET_HEADERS : kind === 'book' ? BOOK_SHEET_HEADERS : GOODS_SHEET_HEADERS,
        rowData: kind === 'magazine' ? buildMagazineSheetRow(product) : kind === 'book' ? buildBookSheetRow(product) : buildGoodsSheetRow(product),
      };
    }),
  };
}

function clearPendingJapaneseTitleMarkers(product) {
  if (!product || typeof product !== 'object') return product;
  return { ...product };
}

async function prepareProductsForSheetSend(items) {
  const sourceItems = Array.isArray(items) ? items : [];
  return sourceItems.map(product => clearPendingJapaneseTitleMarkers(product));
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

  try {
    if (typeof enrichProductWithMangaUpdatesJapaneseTitle === 'function') {
      console.log('[popup] MU enrichment 開始');
      data = await enrichProductWithMangaUpdatesJapaneseTitle(data);
      console.log('[popup] MU enrichment 完了, 日本語タイトル:', data?.日本語タイトル);
    }
  } catch (error) {
    console.warn('[popup] MU enrichment エラー:', error);
  }

  products.push(data);
  save();
  renderList();
  await refreshGasSyncStatus({ updateUi: false, applyMessage: false });
  updateUI();
  if (gasSyncStatus?.state && ['queued', 'running'].includes(gasSyncStatus.state)) {
    startGasSyncStatusPolling();
  }
  setStatus(`✅ ${getProductKindLabel(data)}を追加しました: ${data.商品名?.slice(0, 30) || extractProductCode(data)}`, 'success');
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
  setStatus(
    `✅ 画像:${imageResult.success}/${imageResult.total}件保存 / フォルダ:${imageResult.folders.length} / ${destinationLabel}` +
      (imageResult.failed > 0 ? `（失敗:${imageResult.failed}）` : ''),
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
  setStatus(
    `✅ CSV:${exported}ファイル / 画像:${imageResult.success}/${imageResult.total}件保存 / フォルダ:${imageResult.folders.length} / ${destinationLabel}` +
      (imageResult.failed > 0 ? `（失敗:${imageResult.failed}）` : ''),
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
    setStatus('シート送信用データを準備中...', '');
    const preparedProducts = await prepareProductsForSheetSend(products);
    products = preparedProducts;
    save();
    renderList();
    updateUI();

    setStatus('バックグラウンド送信を登録中...', '');
    const response = await sendRuntimeMessage({
      action: 'queueGasPost',
      url: gasWebAppUrl,
      items: buildGasPayload(preparedProducts).items,
    });

    if (!response?.ok) {
      throw new Error(response?.error || 'バックグラウンド送信の登録に失敗しました');
    }

    clearGasSyncSuccessAutoClearTimer();
    syncGasStatusUi(gasSyncStatus, response.status || null);
    setStatus(gasSyncStatusMessage(gasSyncStatus) || `⏳ ${products.length}件の送信を開始しました`, gasSyncStatusType(gasSyncStatus));
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


initPopup();













