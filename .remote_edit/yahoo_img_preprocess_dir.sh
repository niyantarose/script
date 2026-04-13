#!/usr/bin/env bash
set -Eeuo pipefail

DIR="${1:-}"
if [[ -z "$DIR" ]]; then
  echo "[error] usage: $0 <work_dir>" >&2
  exit 2
fi
if [[ ! -d "$DIR" ]]; then
  echo "[error] dir not found: $DIR" >&2
  exit 2
fi

# 画像変換本体
PREPROCESS_BIN="${PREPROCESS_BIN:-/var/www/html/img/_bin/yahoo_img_preprocess.sh}"
if [[ ! -x "$PREPROCESS_BIN" ]]; then
  echo "[error] missing preprocess bin: $PREPROCESS_BIN" >&2
  exit 2
fi

# notice 注入（GASでやるなら不要）
# NOTICE_MODE:
#   off  = 絶対入れない
#   auto = NOTICE_DIR が存在すれば入れる / 無ければスキップ（失敗にしない）
#   on   = 入れる（ただし無くても失敗にしないようにしてる）
NOTICE_MODE="${NOTICE_MODE:-auto}"
NOTICE_DIR="${NOTICE_DIR:-/var/www/assets/yahoo_notice/tw/goods}"

inject_notice() {
  local mode="$1"
  local notice_dir="$2"

  if [[ "$mode" == "off" ]]; then
    echo "[notice] disabled (NOTICE_MODE=off)"
    return 0
  fi

  if [[ ! -d "$notice_dir" ]]; then
    # ★ここが超重要：無くても exit 1 しない
    echo "[notice] skip (dir not found): $notice_dir"
    return 0
  fi

  shopt -s nullglob
  local files=( "$notice_dir"/* )
  shopt -u nullglob

  if (( ${#files[@]} == 0 )); then
    echo "[notice] skip (empty): $notice_dir"
    return 0
  fi

  # 競合回避しつつコピー
  local copied=0
  for f in "${files[@]}"; do
    [[ -f "$f" ]] || continue
    local base
    base="$(basename "$f")"

    # 同名があれば suffix をつける
    local dst="$DIR/$base"
    if [[ -e "$dst" ]]; then
      local stem="${base%.*}"
      local ext="${base##*.}"
      dst="$DIR/${stem}_notice.${ext}"
      local n=2
      while [[ -e "$dst" ]]; do
        dst="$DIR/${stem}_notice_${n}.${ext}"
        n=$((n+1))
      done
    fi

    cp -f "$f" "$dst"
    copied=$((copied+1))
  done

  echo "[notice] injected: ${copied} file(s) from $notice_dir"
}

# 1) notice 注入（必要なら）
inject_notice "$NOTICE_MODE" "$NOTICE_DIR"

# 2) 画像を jpg 化（ZIPはフラット想定）
#    対象: work_dir 直下のみ
find "$DIR" -maxdepth 1 -type f \
  \( -iname "*.jpg" -o -iname "*.jpeg" -o -iname "*.png" -o -iname "*.webp" -o -iname "*.gif" \) -print0 |
  while IFS= read -r -d '' f; do
    base="$(basename "$f")"
    stem="${base%.*}"
    out="$DIR/${stem}.jpg"

    # 同名上書き時だけ tmp 経由（破壊防止）
    if [[ "$f" == "$out" ]]; then
      tmp="$DIR/.tmp_${stem}_$$.jpg"
      "$PREPROCESS_BIN" "$f" "$tmp"
      mv -f "$tmp" "$out"
    else
      "$PREPROCESS_BIN" "$f" "$out"
      rm -f "$f"
    fi
  done

echo "[ok] preprocess done: $DIR"
exit 0
