import { DEFAULT_JOB_STATE, STORAGE_KEYS } from "./constants.js";

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function storageGet(keys) {
  return chrome.storage.local.get(keys);
}

function storageSet(values) {
  return chrome.storage.local.set(values);
}

function mergeJobState(partial = {}) {
  return {
    ...DEFAULT_JOB_STATE,
    ...partial,
    queue: Array.isArray(partial.queue) ? partial.queue : [],
    errors: Array.isArray(partial.errors) ? partial.errors : []
  };
}

export async function getJobState() {
  const stored = await storageGet(STORAGE_KEYS.jobState);
  return mergeJobState(stored[STORAGE_KEYS.jobState]);
}

export async function setJobState(nextState) {
  const merged = mergeJobState(nextState);
  merged.updatedAt = Date.now();
  await storageSet({ [STORAGE_KEYS.jobState]: merged });
  return merged;
}

export async function patchJobState(mutator) {
  const current = await getJobState();
  const next = mutator(clone(current));
  return setJobState(next);
}
