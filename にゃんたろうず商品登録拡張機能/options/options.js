import { requestSettings, saveSettings } from "../core/messageBus.js";

const form = document.getElementById("settings-form");
const statusEl = document.getElementById("status");

function setStatus(message) {
  statusEl.textContent = message;
}

function applySettings(settings) {
  form.folderNameStrategy.value = settings.folderNameStrategy;
  form.saveMode.value = settings.saveMode;
  form.reRunPolicy.value = settings.reRunPolicy;
  form.saveRawHtml.checked = Boolean(settings.saveRawHtml);
  form.saveExtractedJson.checked = Boolean(settings.saveExtractedJson);
  form.forceJpegConversion.checked = Boolean(settings.forceJpegConversion);
  form["bucket-main"].checked = Boolean(settings.enabledBuckets?.main);
  form["bucket-detail"].checked = Boolean(settings.enabledBuckets?.detail);
  form["bucket-bonus"].checked = Boolean(settings.enabledBuckets?.bonus);
  form["bucket-sample"].checked = Boolean(settings.enabledBuckets?.sample);
}

function collectSettings() {
  return {
    folderNameStrategy: form.folderNameStrategy.value,
    saveMode: form.saveMode.value,
    reRunPolicy: form.reRunPolicy.value,
    saveRawHtml: form.saveRawHtml.checked,
    saveExtractedJson: form.saveExtractedJson.checked,
    forceJpegConversion: form.forceJpegConversion.checked,
    enabledBuckets: {
      main: form["bucket-main"].checked,
      detail: form["bucket-detail"].checked,
      bonus: form["bucket-bonus"].checked,
      sample: form["bucket-sample"].checked
    }
  };
}

async function init() {
  applySettings(await requestSettings());
  setStatus("設定を読み込みました");
}

form.addEventListener("submit", async event => {
  event.preventDefault();
  setStatus("保存中...");
  try {
    const saved = await saveSettings(collectSettings());
    applySettings(saved);
    setStatus("設定を保存しました");
  } catch (error) {
    setStatus(`保存失敗: ${error.message}`);
  }
});

init().catch(error => {
  setStatus(`初期化失敗: ${error.message}`);
});
