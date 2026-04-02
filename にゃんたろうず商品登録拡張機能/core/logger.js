export function logInfo(scope, payload) {
  console.log(`[nyanta:${scope}]`, payload);
}

export function logWarn(scope, payload) {
  console.warn(`[nyanta:${scope}]`, payload);
}

export function logError(scope, error) {
  console.error(`[nyanta:${scope}]`, error);
}
