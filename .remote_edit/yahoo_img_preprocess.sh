#!/usr/bin/env bash
set -euo pipefail

# 画像はアップスケールしない。JPG化だけ行う。
FUZZ="6%"
QUALITY=92
BG="#ffffff"

process_one() {
  local in="$1"
  local out="$2"

  local tmp1
  tmp1="$(dirname "$out")/.tmp_$(basename "$out")_$$.1.jpg"

  # 1) トリム + JPG化（透過は白に）
  convert "$in" \
    -auto-orient \
    -background "$BG" -alpha remove -alpha off \
    -fuzz "$FUZZ" -trim +repage \
    -strip \
    -quality "$QUALITY" \
    "$tmp1"

  # アップスケールは行わず、そのまま出力する。
  mv -f "$tmp1" "$out"
}

if [[ $# -ne 2 ]]; then
  echo "Usage: yahoo_img_preprocess.sh input_image output.jpg" >&2
  exit 2
fi

process_one "$1" "$2"
