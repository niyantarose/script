import { uniqueImageCandidatesByUrl } from "../../../core/dedupe.js";
import { KNOWN_BUCKETS } from "../../../core/constants.js";

export function bucketizeAladinBookImages(raw) {
  const buckets = Object.fromEntries(KNOWN_BUCKETS.map(bucket => [bucket, []]));
  const candidates = uniqueImageCandidatesByUrl(raw.imageCandidates || []);

  for (const candidate of candidates) {
    const bucket = candidate.kindHint === "main" ? "main" : "detail";
    const nextOrder = buckets[bucket].length + 1;
    buckets[bucket].push({
      url: candidate.url,
      order: Number(candidate.orderHint || nextOrder),
      label: `${bucket}_${nextOrder}`,
      sourceSection: candidate.sourceSection || ""
    });
  }

  return buckets;
}
