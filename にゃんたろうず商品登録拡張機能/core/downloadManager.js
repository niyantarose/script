import { buildDataFilename, buildImageFilename, detectImageExtension, resolveConflictAction } from "./filenamePolicy.js";
import { buildDataPath, buildImagePath, resolveFolderNameStrategyValue } from "./pathPolicy.js";
import { enqueueDownloadBatch } from "./downloadQueue.js";
import { buildSequenceScope, nextSequence, resetSequenceScope } from "./sequenceStore.js";
import { uniqueImageCandidatesByUrl } from "./dedupe.js";
import { KNOWN_BUCKETS } from "./constants.js";

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

function textToDataUrl(text, mimeType) {
  const bytes = new TextEncoder().encode(String(text || ""));
  return `data:${mimeType};charset=utf-8;base64,${bytesToBase64(bytes)}`;
}

function jsonToDataUrl(value) {
  return textToDataUrl(JSON.stringify(value, null, 2), "application/json");
}

function buildTask(label, options) {
  return {
    id: `task_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    label,
    options
  };
}

function buildBucketCounts(images) {
  const counts = {};
  for (const bucket of KNOWN_BUCKETS) {
    counts[bucket] = Array.isArray(images[bucket]) ? images[bucket].length : 0;
  }
  return counts;
}

export async function enqueuePayloadDownloads(payload, raw, settings) {
  const tasks = [];
  const conflictAction = resolveConflictAction(settings.reRunPolicy);
  const saveAs = settings.saveMode === "manual";

  payload.product.folderNameStrategyValue = resolveFolderNameStrategyValue(payload, settings);

  const scope = buildSequenceScope({
    folderNameStrategyValue: payload.product.folderNameStrategyValue,
    siteKey: payload.source.siteKey,
    genre: payload.product.genre,
    bucket: "_all"
  });

  if (settings.reRunPolicy !== "continue") {
    await resetSequenceScope(scope);
  }

  for (const bucket of KNOWN_BUCKETS) {
    if (!settings.enabledBuckets?.[bucket]) {
      continue;
    }

    const images = uniqueImageCandidatesByUrl(payload.images?.[bucket] || []);
    if (!images.length) {
      continue;
    }

    for (const image of images) {
      const sequence = await nextSequence(scope);
      const extension = detectImageExtension(image.url, "jpg");
      const filename = buildImageFilename(sequence, extension);
      tasks.push(buildTask(`${bucket}:${filename}`, {
        url: image.url,
        filename: buildImagePath(payload, settings, bucket, filename),
        saveAs,
        conflictAction
      }));
    }
  }

  tasks.unshift(buildTask("product.json", {
    url: jsonToDataUrl(payload),
    filename: buildDataPath(payload, settings, buildDataFilename("product")),
    saveAs,
    conflictAction
  }));

  if (settings.saveExtractedJson) {
    tasks.push(buildTask("extracted.json", {
      url: jsonToDataUrl({
        rawFields: raw.rawFields,
        rawSections: raw.rawSections,
        imageCandidates: raw.imageCandidates,
        pageUrl: raw.pageUrl,
        siteProductCode: raw.siteProductCode
      }),
      filename: buildDataPath(payload, settings, buildDataFilename("extracted")),
      saveAs,
      conflictAction
    }));
  }

  if (settings.saveRawHtml && raw.rawHtml) {
    tasks.push(buildTask("raw.html", {
      url: textToDataUrl(raw.rawHtml, "text/html"),
      filename: buildDataPath(payload, settings, buildDataFilename("raw")),
      saveAs,
      conflictAction
    }));
  }

  const queueResult = await enqueueDownloadBatch(tasks);
  return {
    ...queueResult,
    totalTasks: tasks.length,
    bucketCounts: buildBucketCounts(payload.images),
    folderNameStrategyValue: payload.product.folderNameStrategyValue
  };
}
