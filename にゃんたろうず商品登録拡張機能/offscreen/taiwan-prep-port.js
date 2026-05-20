(function connectTaiwanPrepPort() {
  const port = chrome.runtime.connect({ name: 'taiwanPrep' });

  port.onMessage.addListener(msg => {
    if (!msg || msg.type !== 'taiwanPrepare') return;

    const gasUrl = String(msg.gasUrl || '').trim();
    const prevGas = typeof gasWebAppUrl === 'string' ? gasWebAppUrl : '';

    (async () => {
      try {
        gasWebAppUrl = gasUrl;
        const prepared = await prepareProductsForSheetSend(msg.products || [], { gasUrl });
        const payload = buildGasPayload(prepared);
        port.postMessage({
          type: 'taiwanPrepDone',
          ok: true,
          prepared,
          items: payload.items,
        });
      } catch (error) {
        port.postMessage({
          type: 'taiwanPrepDone',
          ok: false,
          error: error && error.message ? error.message : String(error || 'prepare failed'),
        });
      } finally {
        gasWebAppUrl = prevGas;
      }
    })();
  });
})();
