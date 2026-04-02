chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== 'nyanta:probe-page') {
    return false;
  }

  try {
    const raw = globalThis.__NYANTA_PAGE_PROBE__?.probe(message.adapterId);
    sendResponse({ ok: true, raw });
  } catch (error) {
    sendResponse({ ok: false, error: error?.message || 'probe failed' });
  }

  return true;
});
