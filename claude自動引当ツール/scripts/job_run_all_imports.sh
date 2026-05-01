#!/usr/bin/env bash
set -euo pipefail
cd /home/ubuntu/zaiko-tool/app
exec /usr/bin/flock -n /tmp/zaiko_run_all_imports.lock \
  /home/ubuntu/zaiko-tool/app/venv/bin/python3 /home/ubuntu/zaiko-tool/app/scripts/run_all_imports.py
