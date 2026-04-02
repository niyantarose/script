export const ROOT_DOWNLOAD_FOLDER = "にゃんたろうず商品登録拡張機能";

export const ALARM_NAMES = {
  downloadQueue: "nyanta-download-queue"
};

export const STORAGE_KEYS = {
  settings: "nyanta_settings_v1",
  jobState: "nyanta_job_state_v1",
  sequences: "nyanta_sequences_v1"
};

export const MESSAGE_TYPES = {
  getSettings: "nyanta:get-settings",
  updateSettings: "nyanta:update-settings",
  getPageContext: "nyanta:get-page-context",
  startDownload: "nyanta:start-download",
  getQueueStatus: "nyanta:get-queue-status",
  probePage: "nyanta:probe-page"
};

export const DEFAULT_SETTINGS = {
  folderNameStrategy: "siteProductCode",
  saveMode: "manual",
  forceJpegConversion: false,
  saveRawHtml: false,
  saveExtractedJson: true,
  enabledBuckets: {
    main: true,
    detail: true,
    bonus: true,
    sample: true
  },
  reRunPolicy: "continue"
};

export const DEFAULT_JOB_STATE = {
  queue: [],
  processing: false,
  current: null,
  completed: 0,
  failed: 0,
  errors: [],
  updatedAt: null,
  lastCompleted: null
};

export const KNOWN_BUCKETS = ["main", "detail", "bonus", "sample"];
