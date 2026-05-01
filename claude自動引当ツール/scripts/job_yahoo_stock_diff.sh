#!/usr/bin/env bash
set -euo pipefail
exec /usr/bin/flock -n /tmp/zaiko_yahoo_stock_diff.lock \
  /usr/bin/curl -fsS -X POST "http://127.0.0.1:5000/import/yahoo_stock_diff"
