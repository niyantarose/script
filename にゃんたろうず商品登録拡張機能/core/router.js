import { enqueuePayloadDownloads } from "./downloadManager.js";
import { logError } from "./logger.js";
import { MESSAGE_TYPES } from "./constants.js";
import { getQueueSnapshot } from "./downloadQueue.js";
import { buildFolderPreview } from "./pathPolicy.js";
import { getSettings, updateSettings } from "./settingsStore.js";
import { ensureProbePayload } from "./validators.js";
import { resolveAdapterForUrl } from "../adapters/registry.js";
import { resolveGenreProfile } from "../genres/registry.js";
import { buildPayload } from "../normalizers/buildPayload.js";

let routerRegistered = false;

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs[0] || null;
}

async function collectActiveRaw() {
  const tab = await getActiveTab();
  if (!tab?.id || !tab.url) {
    throw new Error("アクティブタブを取得できませんでした");
  }

  const adapter = resolveAdapterForUrl(tab.url);
  if (!adapter) {
    return { tab, adapter: null, raw: null };
  }

  const raw = ensureProbePayload(await adapter.collectRaw(tab.id));
  return { tab, adapter, raw };
}

async function buildPageContext() {
  const settings = await getSettings();
  const { tab, adapter, raw } = await collectActiveRaw();

  if (!adapter || !raw) {
    return {
      supported: false,
      reason: "現在のページは Phase 1 の対応対象外です",
      url: tab?.url || ""
    };
  }

  const genreProfile = resolveGenreProfile(raw, adapter);
  const payload = buildPayload({ adapter, raw, genreProfile, settings });

  return {
    supported: true,
    adapterId: adapter.id,
    adapterLabel: adapter.label,
    siteKey: payload.source.siteKey,
    genreId: payload.product.genre,
    genreLabel: genreProfile.label,
    title: payload.product.title,
    siteProductCode: payload.product.siteProductCode,
    folderPreview: buildFolderPreview(payload, settings),
    bucketCounts: Object.fromEntries(
      Object.entries(payload.images).map(([bucket, items]) => [bucket, Array.isArray(items) ? items.length : 0])
    ),
    warnings: payload.meta?.warnings || [],
    url: tab.url
  };
}

async function startDownloadForActiveTab() {
  const settings = await getSettings();
  const { adapter, raw } = await collectActiveRaw();

  if (!adapter || !raw) {
    throw new Error("対応サイトのページを開いてください");
  }

  const genreProfile = resolveGenreProfile(raw, adapter);
  const payload = buildPayload({ adapter, raw, genreProfile, settings });
  const queueResult = await enqueuePayloadDownloads(payload, raw, settings);

  return {
    title: payload.product.title,
    siteKey: payload.source.siteKey,
    genre: payload.product.genre,
    folderPreview: buildFolderPreview(payload, settings),
    queueResult
  };
}

async function routeMessage(message) {
  switch (message?.type) {
    case MESSAGE_TYPES.getSettings:
      return { ok: true, settings: await getSettings() };

    case MESSAGE_TYPES.updateSettings:
      return { ok: true, settings: await updateSettings(message.settings || {}) };

    case MESSAGE_TYPES.getPageContext:
      return { ok: true, context: await buildPageContext() };

    case MESSAGE_TYPES.startDownload:
      return { ok: true, result: await startDownloadForActiveTab() };

    case MESSAGE_TYPES.getQueueStatus:
      return { ok: true, state: await getQueueSnapshot() };

    default:
      return { ok: false, error: "unknown message" };
  }
}

export function registerRouter() {
  if (routerRegistered) {
    return;
  }

  routerRegistered = true;

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    routeMessage(message)
      .then(sendResponse)
      .catch(error => {
        logError("router", error);
        sendResponse({ ok: false, error: error?.message || "unexpected error" });
      });
    return true;
  });
}
