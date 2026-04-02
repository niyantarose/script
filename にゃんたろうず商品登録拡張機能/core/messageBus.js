import { MESSAGE_TYPES } from "./constants.js";

export function sendRuntimeMessage(type, payload = {}) {
  return chrome.runtime.sendMessage({ type, ...payload });
}

export async function requestSettings() {
  const response = await sendRuntimeMessage(MESSAGE_TYPES.getSettings);
  if (!response?.ok) {
    throw new Error(response?.error || "settings read failed");
  }
  return response.settings;
}

export async function saveSettings(settings) {
  const response = await sendRuntimeMessage(MESSAGE_TYPES.updateSettings, { settings });
  if (!response?.ok) {
    throw new Error(response?.error || "settings save failed");
  }
  return response.settings;
}

export async function requestPageContext() {
  const response = await sendRuntimeMessage(MESSAGE_TYPES.getPageContext);
  if (!response?.ok) {
    throw new Error(response?.error || "page context failed");
  }
  return response.context;
}

export async function requestQueueStatus() {
  const response = await sendRuntimeMessage(MESSAGE_TYPES.getQueueStatus);
  if (!response?.ok) {
    throw new Error(response?.error || "queue status failed");
  }
  return response.state;
}

export async function startDownloadRun() {
  const response = await sendRuntimeMessage(MESSAGE_TYPES.startDownload);
  if (!response?.ok) {
    throw new Error(response?.error || "download start failed");
  }
  return response.result;
}
