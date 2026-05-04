#!/usr/bin/env bash
set -euo pipefail
exec /usr/bin/flock -n /tmp/zaiko_yahoo_stock_full.lock \
  /home/ubuntu/zaiko-tool/app/venv/bin/python \
  /home/ubuntu/zaiko-tool/app/scripts/import_yahoo_stock_full.py
