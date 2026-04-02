(async () => {
  const status = document.getElementById('router-status');

  function go(path) {
    location.replace(path);
  }

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const url = String(tab?.url || '');

    if (/^https:\/\/www\.books\.com\.tw\/products\//i.test(url)) {
      go('./taiwan/popup.html');
      return;
    }

    if (/^https:\/\/www\.aladin\.co\.kr\/shop\/wproduct\.aspx/i.test(url)) {
      go('./aladin/popup.html');
      return;
    }

    go('./home.html');
  } catch (error) {
    status.textContent = error?.message || 'popup route failed';
    go('./home.html');
  }
})();
