import { ALARM_NAMES } from "./constants.js";
import { getJobState, patchJobState } from "./jobStore.js";
import { logError } from "./logger.js";

let isProcessing = false;
let listenersRegistered = false;

function executeDownload(task) {
  return new Promise((resolve, reject) => {
    chrome.downloads.download(task.options, downloadId => {
      const runtimeError = chrome.runtime.lastError;
      if (runtimeError || !downloadId) {
        reject(new Error(runtimeError?.message || "download failed"));
        return;
      }
      resolve(downloadId);
    });
  });
}

export function registerDownloadQueue() {
  if (listenersRegistered) {
    return;
  }

  listenersRegistered = true;

  chrome.alarms.onAlarm.addListener(alarm => {
    if (alarm?.name === ALARM_NAMES.downloadQueue) {
      void processDownloadQueue();
    }
  });

  chrome.runtime.onStartup?.addListener(() => {
    void processDownloadQueue();
  });

  chrome.runtime.onInstalled?.addListener(() => {
    void processDownloadQueue();
  });
}

export async function enqueueDownloadBatch(tasks) {
  if (!Array.isArray(tasks) || tasks.length === 0) {
    return { accepted: 0 };
  }

  await patchJobState(state => {
    state.queue = [...state.queue, ...tasks];
    return state;
  });

  chrome.alarms.create(ALARM_NAMES.downloadQueue, { when: Date.now() + 50 });
  void processDownloadQueue();
  return { accepted: tasks.length };
}

export async function getQueueSnapshot() {
  return getJobState();
}

export async function processDownloadQueue() {
  if (isProcessing) {
    return;
  }

  isProcessing = true;

  try {
    while (true) {
      const state = await getJobState();
      const task = state.queue[0];

      if (!task) {
        await patchJobState(current => {
          current.processing = false;
          current.current = null;
          return current;
        });
        break;
      }

      await patchJobState(current => {
        current.processing = true;
        current.current = {
          id: task.id,
          label: task.label,
          filename: task.options.filename
        };
        return current;
      });

      try {
        const downloadId = await executeDownload(task);
        await patchJobState(current => {
          current.queue = current.queue.slice(1);
          current.processing = false;
          current.current = null;
          current.completed += 1;
          current.lastCompleted = {
            id: task.id,
            label: task.label,
            filename: task.options.filename,
            downloadId,
            completedAt: new Date().toISOString()
          };
          return current;
        });
      } catch (error) {
        logError("downloadQueue", error);
        await patchJobState(current => {
          current.queue = current.queue.slice(1);
          current.processing = false;
          current.current = null;
          current.failed += 1;
          current.errors = [
            {
              id: task.id,
              label: task.label,
              filename: task.options.filename,
              message: error?.message || "download failed",
              failedAt: new Date().toISOString()
            },
            ...current.errors
          ].slice(0, 20);
          return current;
        });
      }
    }
  } finally {
    isProcessing = false;
  }
}
