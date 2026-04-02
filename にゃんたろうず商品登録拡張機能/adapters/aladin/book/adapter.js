import { MESSAGE_TYPES } from "../../../core/constants.js";
import { normalizeAladinBookRaw } from "./rawExtract.js";
import { bucketizeAladinBookImages } from "./imageExtract.js";

function requestProbe(tabId, adapterId) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, {
      type: MESSAGE_TYPES.probePage,
      adapterId
    }, response => {
      const runtimeError = chrome.runtime.lastError;
      if (runtimeError) {
        reject(new Error(runtimeError.message || "probe request failed"));
        return;
      }
      if (!response?.ok) {
        reject(new Error(response?.error || "probe response missing"));
        return;
      }
      resolve(response.raw);
    });
  });
}

export const aladinBookAdapter = {
  id: "aladin_book",
  siteKey: "aladin_book",
  label: "Aladin Book",
  country: "aladin",
  match(url) {
    return /^https:\/\/www\.aladin\.co\.kr\/shop\/wproduct\.aspx/i.test(String(url || ""));
  },
  async collectRaw(tabId) {
    return normalizeAladinBookRaw(await requestProbe(tabId, this.id));
  },
  buildImageBuckets(raw) {
    return bucketizeAladinBookImages(raw);
  }
};
