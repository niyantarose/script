import { DEFAULT_SETTINGS, STORAGE_KEYS } from "./constants.js";

function storageGet(keys) {
  return chrome.storage.local.get(keys);
}

function storageSet(values) {
  return chrome.storage.local.set(values);
}

function mergeSettings(partial = {}) {
  const enabledBuckets = {
    ...DEFAULT_SETTINGS.enabledBuckets,
    ...(partial.enabledBuckets || {})
  };

  return {
    ...DEFAULT_SETTINGS,
    ...partial,
    enabledBuckets
  };
}

export async function getSettings() {
  const stored = await storageGet(STORAGE_KEYS.settings);
  return mergeSettings(stored[STORAGE_KEYS.settings]);
}

export async function updateSettings(nextPartial = {}) {
  const current = await getSettings();
  const next = mergeSettings({
    ...current,
    ...nextPartial,
    enabledBuckets: {
      ...current.enabledBuckets,
      ...(nextPartial.enabledBuckets || {})
    }
  });
  await storageSet({ [STORAGE_KEYS.settings]: next });
  return next;
}
