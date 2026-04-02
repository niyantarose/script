import { STORAGE_KEYS } from "./constants.js";

function storageGet(keys) {
  return chrome.storage.local.get(keys);
}

function storageSet(values) {
  return chrome.storage.local.set(values);
}

async function getSequenceMap() {
  const stored = await storageGet(STORAGE_KEYS.sequences);
  return stored[STORAGE_KEYS.sequences] || {};
}

async function setSequenceMap(sequenceMap) {
  await storageSet({ [STORAGE_KEYS.sequences]: sequenceMap });
  return sequenceMap;
}

export function buildSequenceScope({ folderNameStrategyValue, siteKey, genre, bucket }) {
  return [folderNameStrategyValue, siteKey, genre, bucket].join("::");
}

export async function resetSequenceScope(scope) {
  const sequenceMap = await getSequenceMap();
  sequenceMap[scope] = 0;
  await setSequenceMap(sequenceMap);
}

export async function nextSequence(scope) {
  const sequenceMap = await getSequenceMap();
  const nextValue = Number(sequenceMap[scope] || 0) + 1;
  sequenceMap[scope] = nextValue;
  await setSequenceMap(sequenceMap);
  return nextValue;
}
