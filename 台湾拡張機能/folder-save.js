// folder-save.js - File System Access API でフォルダに直接保存する共通ユーティリティ
// showDirectoryPicker() で選んだフォルダを IndexedDB に記憶し、
// 商品フォルダ作成 + 連番画像保存を行う。

const FOLDER_DB_NAME = 'booksTwFolderHandleDB';
const FOLDER_STORE_NAME = 'handles';
const FOLDER_HANDLE_KEY = 'downloadRoot';
const FOLDER_STATUS_KEY = 'booksTwFolderPath';

function _openFolderDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(FOLDER_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(FOLDER_STORE_NAME)) {
        db.createObjectStore(FOLDER_STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function _idbSet(key, value) {
  const db = await _openFolderDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(FOLDER_STORE_NAME, 'readwrite');
    tx.objectStore(FOLDER_STORE_NAME).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function _idbGet(key) {
  const db = await _openFolderDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(FOLDER_STORE_NAME, 'readonly');
    const req = tx.objectStore(FOLDER_STORE_NAME).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// フォルダ選択ダイアログを出して IndexedDB に保存
async function pickAndSaveFolder() {
  const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
  await _idbSet(FOLDER_HANDLE_KEY, dirHandle);
  // フォルダ名を chrome.storage にも保存（UI表示用）
  await new Promise(resolve => {
    chrome.storage.local.set({ [FOLDER_STATUS_KEY]: dirHandle.name }, resolve);
  });
  return dirHandle.name;
}

// 保存済みフォルダハンドルを取得（なければ null）
async function getSavedFolderHandle() {
  try {
    return await _idbGet(FOLDER_HANDLE_KEY) || null;
  } catch {
    return null;
  }
}

// 保存済みフォルダ名を取得（UI表示用）
async function getSavedFolderName() {
  return new Promise(resolve => {
    chrome.storage.local.get(FOLDER_STATUS_KEY, result => {
      resolve(result[FOLDER_STATUS_KEY] || '');
    });
  });
}

// フォルダハンドルの書き込み権限を確認・リクエスト
async function ensureFolderPermission(handle) {
  const opts = { mode: 'readwrite' };
  if ((await handle.queryPermission(opts)) === 'granted') return true;
  if ((await handle.requestPermission(opts)) === 'granted') return true;
  return false;
}

function _sanitizeSegment(text) {
  return String(text || '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '') || 'item';
}

function _guessExt(url) {
  try {
    const ext = new URL(url).pathname.split('.').pop()?.toLowerCase() || '';
    if (ext === 'jpeg') return 'jpg';
    if (['jpg', 'png', 'webp', 'gif'].includes(ext)) return ext;
  } catch { /* ignore */ }
  return 'jpg';
}

// 画像を fetch して Blob として返す
async function _fetchImageBlob(imageUrl) {
  const response = await fetch(imageUrl, { credentials: 'omit' });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${imageUrl}`);
  return await response.blob();
}

// 商品フォルダに画像を連番保存する
// rootHandle: FileSystemDirectoryHandle（選択済みのルートフォルダ）
// folderName: サブフォルダ名（商品ID or タイトル）
// imageUrls: 画像URL配列
// onProgress: (seq, total, ok) => void  進捗コールバック
async function saveImagesToFolder(rootHandle, folderName, imageUrls, onProgress) {
  const safeName = _sanitizeSegment(folderName);
  const productDir = await rootHandle.getDirectoryHandle(safeName, { create: true });

  let success = 0;
  let failed = 0;
  const total = imageUrls.length;

  for (let i = 0; i < imageUrls.length; i++) {
    const seq = i + 1;
    const ext = _guessExt(imageUrls[i]);
    try {
      const blob = await _fetchImageBlob(imageUrls[i]);
      const fileHandle = await productDir.getFileHandle(`${seq}.${ext}`, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      success++;
      if (onProgress) onProgress(seq, total, true);
    } catch (err) {
      console.error(`画像保存失敗 [${seq}/${total}]: ${imageUrls[i]}`, err);
      failed++;
      if (onProgress) onProgress(seq, total, false);
    }
  }

  return { folder: safeName, total, success, failed };
}
