// folder-save.js - File System Access API で選択フォルダへ直接保存する共通ユーティリティ
// 商品フォルダ作成、連番画像保存、テキストファイル保存を行う。

const FOLDER_DB_NAME = 'aladinFolderHandleDB';
const FOLDER_STORE_NAME = 'handles';
const FOLDER_HANDLE_KEY = 'downloadRoot';
const FOLDER_STATUS_KEY = 'aladinFolderPath';

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

async function pickAndSaveFolder() {
  const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads', id: 'nyantarose-download-root' });
  await _idbSet(FOLDER_HANDLE_KEY, dirHandle);
  await new Promise(resolve => {
    chrome.storage.local.set({ [FOLDER_STATUS_KEY]: dirHandle.name }, resolve);
  });
  return dirHandle.name;
}

async function getSavedFolderHandle() {
  try {
    return await _idbGet(FOLDER_HANDLE_KEY) || null;
  } catch {
    return null;
  }
}

async function getSavedFolderName() {
  return new Promise(resolve => {
    chrome.storage.local.get(FOLDER_STATUS_KEY, result => {
      resolve(result[FOLDER_STATUS_KEY] || '');
    });
  });
}

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

function _splitRelativePath(relativePath) {
  return String(relativePath || '')
    .split(/[\\/]+/)
    .map(segment => _sanitizeSegment(segment))
    .filter(Boolean);
}

async function getOrCreateDirectoryPath(rootHandle, relativePath) {
  let current = rootHandle;
  for (const segment of _splitRelativePath(relativePath)) {
    current = await current.getDirectoryHandle(segment, { create: true });
  }
  return current;
}

function _splitFileName(fileName) {
  const safeName = _sanitizeSegment(fileName || 'file');
  const match = /^(.*?)(\.[^.]+)?$/.exec(safeName) || [];
  return {
    base: match[1] || 'file',
    ext: match[2] || '',
    full: safeName,
  };
}

async function _fileExists(dirHandle, fileName) {
  try {
    await dirHandle.getFileHandle(fileName);
    return true;
  } catch (error) {
    if (error?.name === 'NotFoundError') return false;
    throw error;
  }
}

async function getUniqueFileName(dirHandle, preferredName) {
  const parts = _splitFileName(preferredName);
  if (!(await _fileExists(dirHandle, parts.full))) {
    return parts.full;
  }

  for (let index = 1; index < 10000; index += 1) {
    const candidate = `${parts.base} (${index})${parts.ext}`;
    if (!(await _fileExists(dirHandle, candidate))) {
      return candidate;
    }
  }

  throw new Error(`重複ファイル名を解決できません: ${preferredName}`);
}

async function getNextImageSequence(dirHandle) {
  let max = 0;
  for await (const entry of dirHandle.values()) {
    if (entry.kind !== 'file') continue;
    const match = /^(\d+)(?:\.[^.]+)?$/i.exec(entry.name || '');
    if (!match) continue;
    const value = Number.parseInt(match[1], 10);
    if (Number.isFinite(value)) {
      max = Math.max(max, value);
    }
  }
  return max + 1;
}

function _guessExt(url, blobType = '') {
  const type = String(blobType || '').toLowerCase();
  if (type === 'image/jpeg' || type === 'image/jpg') return 'jpg';
  if (type === 'image/png') return 'png';
  if (type === 'image/webp') return 'webp';
  if (type === 'image/gif') return 'gif';

  try {
    const ext = new URL(url).pathname.split('.').pop()?.toLowerCase() || '';
    if (ext === 'jpeg') return 'jpg';
    if (['jpg', 'png', 'webp', 'gif'].includes(ext)) return ext;
  } catch {
    // ignore
  }
  return 'jpg';
}

async function _fetchImageBlob(imageUrl) {
  const response = await fetch(imageUrl, {
    credentials: 'omit',
    cache: 'no-store',
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${imageUrl}`);
  }
  const blob = await response.blob();
  if (!blob || !blob.size) {
    throw new Error(`empty blob: ${imageUrl}`);
  }
  return blob;
}

async function _convertBlobToJpeg(blob) {
  if (/^image\/jpe?g$/i.test(String(blob?.type || ''))) {
    return blob;
  }

  let bitmap = null;
  try {
    bitmap = await createImageBitmap(blob);
    const width = Math.max(1, bitmap.width || 1);
    const height = Math.max(1, bitmap.height || 1);

    let canvas;
    let context;
    if (typeof OffscreenCanvas !== 'undefined') {
      canvas = new OffscreenCanvas(width, height);
      context = canvas.getContext('2d', { alpha: false });
    } else {
      canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      context = canvas.getContext('2d', { alpha: false });
    }

    if (!context) {
      throw new Error('2D context unavailable');
    }

    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);
    context.drawImage(bitmap, 0, 0, width, height);

    if (typeof canvas.convertToBlob === 'function') {
      return await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.92 });
    }

    return await new Promise((resolve, reject) => {
      canvas.toBlob(result => {
        if (result) {
          resolve(result);
          return;
        }
        reject(new Error('canvas.toBlob failed'));
      }, 'image/jpeg', 0.92);
    });
  } catch (error) {
    throw new Error(`JPEG変換失敗: ${error?.message || error}`);
  } finally {
    if (bitmap && typeof bitmap.close === 'function') {
      try { bitmap.close(); } catch { /* ignore */ }
    }
  }
}

async function saveTextFileToFolder(rootHandle, relativeFilePath, text, options = {}) {
  const segments = _splitRelativePath(relativeFilePath);
  if (!segments.length) {
    throw new Error('保存ファイルパスが空です');
  }

  const fileName = segments.pop();
  const dirHandle = await getOrCreateDirectoryPath(rootHandle, segments.join('/'));
  const resolvedName = options.uniquify === false ? _splitFileName(fileName).full : await getUniqueFileName(dirHandle, fileName);
  const fileHandle = await dirHandle.getFileHandle(resolvedName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(String(text ?? ''));
  await writable.close();

  const resolvedPath = [...segments, resolvedName].join('/');
  return {
    fileName: resolvedName,
    path: resolvedPath,
  };
}

async function saveImagesToFolder(rootHandle, relativeFolderPath, imageUrls, onProgress, options = {}) {
  const dirHandle = await getOrCreateDirectoryPath(rootHandle, relativeFolderPath);
  const forceJpeg = options.forceJpeg !== false;
  const urls = [];
  const seen = new Set();

  for (const rawUrl of Array.isArray(imageUrls) ? imageUrls : []) {
    const url = String(rawUrl || '').trim();
    if (!/^https?:\/\//i.test(url)) continue;
    if (seen.has(url)) continue;
    seen.add(url);
    urls.push(url);
  }

  let nextSequence = await getNextImageSequence(dirHandle);
  let success = 0;
  let failed = 0;

  for (let index = 0; index < urls.length; index += 1) {
    const imageUrl = urls[index];
    const seq = nextSequence;
    nextSequence += 1;

    try {
      const originalBlob = await _fetchImageBlob(imageUrl);
      const blob = forceJpeg ? await _convertBlobToJpeg(originalBlob) : originalBlob;
      const ext = forceJpeg ? 'jpg' : _guessExt(imageUrl, blob.type);
      const fileHandle = await dirHandle.getFileHandle(`${seq}.${ext}`, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      success += 1;
      if (onProgress) onProgress(seq, urls.length, true, imageUrl);
    } catch (error) {
      console.error(`画像保存失敗 [${seq}/${urls.length}]`, imageUrl, error);
      failed += 1;
      if (onProgress) onProgress(seq, urls.length, false, imageUrl);
    }
  }

  return {
    folder: _splitRelativePath(relativeFolderPath).join('/'),
    total: urls.length,
    success,
    failed,
    nextSequence,
  };
}

