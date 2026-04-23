#!/usr/bin/env python3
"""高速なローカル画像分割ツール。

GAS の実行制限を避け、ローカルPC上で高速に画像を分割保存するための CLI。
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable, List

from PIL import Image

Image.MAX_IMAGE_PIXELS = None


def parse_cut_points(raw: str) -> List[int]:
    if not raw:
        return []
    points = sorted({int(p.strip()) for p in raw.split(",") if p.strip()})
    if any(p <= 0 for p in points):
        raise ValueError("cut points は 1 以上の整数で指定してください")
    return points


def build_segments(height: int, cut_points: List[int]) -> List[tuple[int, int]]:
    usable = [p for p in cut_points if p < height]
    if cut_points and not usable:
        raise ValueError(f"指定された cut points が画像の高さ({height}px)を超えています")

    boundaries = [0, *usable, height]
    segments: List[tuple[int, int]] = []
    for y0, y1 in zip(boundaries, boundaries[1:]):
        if y1 <= y0:
            continue
        segments.append((y0, y1))
    return segments


def build_segments_by_height(height: int, split_height: int, overlap: int) -> List[tuple[int, int]]:
    if split_height <= 0:
        raise ValueError("--split-height は 1 以上を指定してください")
    if overlap < 0:
        raise ValueError("--overlap は 0 以上を指定してください")
    if overlap >= split_height:
        raise ValueError("--overlap は --split-height より小さくしてください")

    segments: List[tuple[int, int]] = []
    y0 = 0
    step = split_height - overlap
    while y0 < height:
        y1 = min(y0 + split_height, height)
        segments.append((y0, y1))
        if y1 == height:
            break
        y0 += step
    return segments


def save_split_images(
    image_path: Path,
    output_dir: Path,
    cut_points: List[int] | None,
    split_height: int | None,
    overlap: int,
    output_format: str,
    jpeg_quality: int,
) -> list[Path]:
    with Image.open(image_path) as im:
        src = im.convert("RGB") if output_format in {"jpg", "jpeg"} else im.copy()
        width, height = src.size

        if cut_points is not None:
            segments = build_segments(height, cut_points)
        else:
            assert split_height is not None
            segments = build_segments_by_height(height, split_height, overlap)

        base = image_path.stem
        ext = "jpg" if output_format == "jpeg" else output_format

        saved_paths: list[Path] = []
        for idx, (y0, y1) in enumerate(segments, start=1):
            cropped = src.crop((0, y0, width, y1))
            out_path = output_dir / f"{base}_part{idx:03d}_{y0}-{y1}.{ext}"

            save_kwargs = {}
            if ext == "jpg":
                save_kwargs = {"quality": jpeg_quality, "optimize": True}
            cropped.save(out_path, **save_kwargs)
            saved_paths.append(out_path)

    return saved_paths


def iter_images(paths: Iterable[Path]) -> list[Path]:
    out: list[Path] = []
    for p in paths:
        if p.is_dir():
            for ext in ("*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp", "*.tif", "*.tiff"):
                out.extend(sorted(p.glob(ext)))
        else:
            out.append(p)
    unique_sorted = sorted({p.resolve() for p in out if p.exists()})
    return unique_sorted


def main() -> None:
    parser = argparse.ArgumentParser(description="高速ローカル画像分割ツール")
    parser.add_argument("inputs", nargs="+", type=Path, help="画像ファイルまたはフォルダ")
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("output"), help="出力先フォルダ")

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--cuts", type=str, help="分割Y座標(カンマ区切り)。例: 1200,2400,3600")
    mode.add_argument("--split-height", type=int, help="固定高さで自動分割")

    parser.add_argument("--overlap", type=int, default=0, help="自動分割時の重なり(px)")
    parser.add_argument("--format", choices=["png", "jpg", "jpeg", "webp"], default="png", help="出力形式")
    parser.add_argument("--jpeg-quality", type=int, default=92, help="JPEG品質(1-100)")
    parser.add_argument("--workers", type=int, default=4, help="並列処理数")

    args = parser.parse_args()

    if not 1 <= args.jpeg_quality <= 100:
        raise ValueError("--jpeg-quality は 1-100 で指定してください")

    output_dir: Path = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    images = iter_images(args.inputs)
    if not images:
        raise FileNotFoundError("入力画像が見つかりませんでした")

    cut_points = parse_cut_points(args.cuts) if args.cuts else None

    print(f"入力画像数: {len(images)}")
    print(f"出力先: {output_dir.resolve()}")

    futures = []
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        for image_path in images:
            futures.append(
                executor.submit(
                    save_split_images,
                    image_path=image_path,
                    output_dir=output_dir,
                    cut_points=cut_points,
                    split_height=args.split_height,
                    overlap=args.overlap,
                    output_format=args.format,
                    jpeg_quality=args.jpeg_quality,
                )
            )

        total = 0
        for fut in as_completed(futures):
            saved = fut.result()
            total += len(saved)
            if saved:
                print(f"保存完了: {saved[0].parent.name} / {saved[0].stem.split('_part')[0]} ({len(saved)}枚)")

    print(f"\n完了: 合計 {total} ファイル保存しました")


if __name__ == "__main__":
    main()
