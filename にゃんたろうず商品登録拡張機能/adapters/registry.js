import { aladinBookAdapter } from "./aladin/book/adapter.js";

const ADAPTERS = [aladinBookAdapter];

export function getRegisteredAdapters() {
  return ADAPTERS.slice();
}

export function resolveAdapterForUrl(url) {
  return ADAPTERS.find(adapter => adapter.match(url)) || null;
}
