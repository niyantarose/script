#!/usr/bin/env bash
set -euo pipefail
exec /usr/bin/flock -n /tmp/zaiko_yahoo_orders.lock \
  /usr/bin/curl -fsS --max-time 900 -X POST "http://127.0.0.1:5000/import/yahoo_orders" \
    -H "Content-Type: application/json" \
    -d '{"days":30}'
