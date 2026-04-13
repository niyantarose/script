#!/usr/bin/env bash
set -u

FORCE=0
if [[ "${1:-}" == "--force" ]]; then
  FORCE=1
  shift
fi

TSV="${1:-}"
if [[ -z "$TSV" || ! -f "$TSV" ]]; then
  echo "Usage: $0 [--force] /path/to/list.tsv" >&2
  exit 2
fi

ROOT="/var/www/html/img"
TMPDIR="/tmp/imgcv_work"
mkdir -p "$TMPDIR"

# ImageMagick: /usr/bin/convert or /usr/bin/magick
IM_BIN=""
if command -v magick >/dev/null 2>&1; then
  IM_BIN="magick"
elif command -v convert >/dev/null 2>&1; then
  IM_BIN="convert"
fi

# webp decode tool (optional)
DWEBP_BIN=""
if command -v dwebp >/dev/null 2>&1; then
  DWEBP_BIN="dwebp"
fi

# ===== リサイズ方針 =====
# 小さい画像はそのまま維持し、大きすぎる画像だけ縮小する
MAX_W=2000
MAX_H=2000

# ImageMagick 呼び出し差分吸収
im_run() {
  if [[ "$IM_BIN" == "magick" ]]; then
    magick "$@"
  else
    convert "$@"
  fi
}

# 1行: URL \t CODE \t NAME(jpg想定)
while IFS=$'\t' read -r URL CODE NAME; do
  URL="${URL%$'\r'}"
  CODE="${CODE%$'\r'}"
  NAME="${NAME%$'\r'}"

  [[ -z "$URL" || -z "$CODE" || -z "$NAME" ]] && continue

  OUTDIR="${ROOT}/${CODE}"
  OUT="${OUTDIR}/${NAME}"
  mkdir -p "$OUTDIR"

  # 既に存在してて強制じゃないならキャッシュ扱い
  if [[ -f "$OUT" && "$FORCE" -ne 1 ]]; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tOK\tCACHED"
    continue
  fi

  # ダウンロード先（まずは生データ）
  INBIN="${TMPDIR}/in_${CODE}_$$.bin"
  rm -f "$INBIN"

  # ダウンロード（リダイレクト追従、UA固定）
  if ! curl -sS -L -A 'Mozilla/5.0 (IMGCV)' -o "$INBIN" "$URL"; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tFAIL\tCURL_ERROR"
    rm -f "$INBIN"
    continue
  fi

  # 小さすぎるのは弾く（HTMLやエラーの可能性）
  SIZE=$(stat -c%s "$INBIN" 2>/dev/null || echo 0)
  if [[ "$SIZE" -lt 1024 ]]; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tFAIL\tTOO_SMALL"
    rm -f "$INBIN"
    continue
  fi

  # Content判定
  FTYPE=$(file -b "$INBIN" 2>/dev/null || echo "")

  # HTML等を弾く
  if echo "$FTYPE" | grep -qiE 'HTML|XML|JSON|text'; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tFAIL\tNOT_IMAGE(${FTYPE})"
    rm -f "$INBIN"
    continue
  fi

  # 出力は「NAMEがjpgでも実体がwebpでも」JPGにする（Yahoo用）
  OUTTMP="${TMPDIR}/out_${CODE}_$$.jpg"
  rm -f "$OUTTMP"

  # 変換（優先：ImageMagick）
  CONVERT_OK=0
  if [[ -n "$IM_BIN" ]]; then
    # IMがWebP読めるならこれでOK
    if im_run "$INBIN" -strip -quality 92 "$OUTTMP" >/dev/null 2>&1; then
      CONVERT_OK=1
    fi
  fi

  # IMでダメなら dwebp → convert/magick にフォールバック
  if [[ "$CONVERT_OK" -ne 1 && -n "$DWEBP_BIN" && -n "$IM_BIN" ]]; then
    PNGTMP="${TMPDIR}/out_${CODE}_$$.png"
    rm -f "$PNGTMP"
    if $DWEBP_BIN "$INBIN" -o "$PNGTMP" >/dev/null 2>&1; then
      if im_run "$PNGTMP" -strip -quality 92 "$OUTTMP" >/dev/null 2>&1; then
        CONVERT_OK=1
      fi
    fi
    rm -f "$PNGTMP"
  fi

  if [[ "$CONVERT_OK" -ne 1 ]]; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tFAIL\tCONVERT_FAILED(${FTYPE})"
    rm -f "$INBIN" "$OUTTMP"
    continue
  fi

  # ===== 解像度チェック & リサイズ =====
  # 小さい画像はそのまま維持し、大きすぎる画像だけ縮小する
  if [[ -n "$IM_BIN" ]]; then
    W=$(im_run "$OUTTMP" -ping -format "%w" info: 2>/dev/null || echo 0)
    H=$(im_run "$OUTTMP" -ping -format "%h" info: 2>/dev/null || echo 0)

    if [[ "$W" -gt 0 && "$H" -gt 0 ]]; then
      RESIZE_FLAG="NONE"

      # 大きすぎる場合だけ、上限内に収まるよう縮小する
      if [[ "$W" -gt "$MAX_W" || "$H" -gt "$MAX_H" ]]; then
        RESIZE_FLAG="DOWNSCALE"
        im_run "$OUTTMP" \
          -filter Lanczos -resize "${MAX_W}x${MAX_H}>" \
          -strip -quality 92 \
          "$OUTTMP" >/dev/null 2>&1 || true
      fi

      # ログ（RESIZED のときだけOKログを上書きで出す）
      if [[ "$RESIZE_FLAG" != "NONE" ]]; then
        W2=$(im_run "$OUTTMP" -ping -format "%w" info: 2>/dev/null || echo 0)
        H2=$(im_run "$OUTTMP" -ping -format "%h" info: 2>/dev/null || echo 0)
        echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tOK\tRESIZED:${RESIZE_FLAG} ${W}x${H}=>${W2}x${H2}"
      fi
    fi
  fi

  # JPEG確認
  if ! file "$OUTTMP" | grep -qi "JPEG"; then
    echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tFAIL\tNOT_JPEG"
    rm -f "$INBIN" "$OUTTMP"
    continue
  fi

  # 配置（www-data運用なら所有権合わせ）
  mv -f "$OUTTMP" "$OUT"
  chmod 664 "$OUT" >/dev/null 2>&1 || true

  # ここまでで「RESIZEDログ」を出していない場合は通常OKログ
  echo -e "IMGCV_RESULT\t${CODE}\t${NAME}\tOK\tDOWNLOADED"
  rm -f "$INBIN"

done < "$TSV"

echo "DONE_IMGCV"
