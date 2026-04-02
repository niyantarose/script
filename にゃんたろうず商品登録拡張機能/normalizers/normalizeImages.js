import { KNOWN_BUCKETS } from "../core/constants.js";

export function normalizeImages({ adapter, raw, genreProfile }) {
  const bucketed = adapter.buildImageBuckets ? adapter.buildImageBuckets(raw) : {};
  const allowedBuckets = new Set(genreProfile.imageBuckets || []);

  return Object.fromEntries(
    KNOWN_BUCKETS.map(bucket => [
      bucket,
      allowedBuckets.has(bucket) ? (bucketed[bucket] || []) : []
    ])
  );
}
