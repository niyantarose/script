#!/usr/bin/env bash
set -euo pipefail
exec /usr/bin/flock -n /tmp/zaiko_yahoo_stock_full.lock \
  /usr/bin/curl -fsS --max-time 1800 -X POST "http://127.0.0.1:5000/import/yahoo_stock"
