#!/usr/bin/env python3
from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import (
    END,
    BOTH,
    LEFT,
    RIGHT,
    VERTICAL,
    StringVar,
    Tk,
    filedialog,
    messagebox,
    ttk,
)
from tkinter.scrolledtext import ScrolledText

import requests
from PIL import Image

Image.MAX_IMAGE_PIXELS = None

APP_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = APP_DIR / "settings.json"


@dataclass
class SplitOptions:
    mode: str
    cuts: list[int]
    split_height: int
    overlap: int
    output_format: str
    jpeg_quality: int
    workers: int


def parse_cuts(raw: str) -> list[int]:
    if not raw.strip():
        return []
    values = sorted({int(v.strip()) for v in raw.split(",") if v.strip()})
    if any(v <= 0 for v in values):
        raise ValueError("cuts は 1 以上の整数で指定してください")
    return values


def split_by_cuts(height: int, cuts: list[int]) -> list[tuple[int, int]]:
    filtered = [c for c in cuts if 0 < c < height]
    boundaries = [0, *filtered, height]
    return [(y0, y1) for y0, y1 in zip(boundaries, boundaries[1:]) if y1 > y0]


def split_by_height(height: int, split_height: int, overlap: int) -> list[tuple[int, int]]:
    if split_height <= 0:
        raise ValueError("split height は 1 以上")
    if overlap < 0 or overlap >= split_height:
        raise ValueError("overlap は 0以上 かつ split height 未満")

    result: list[tuple[int, int]] = []
    step = split_height - overlap
    y = 0
    while y < height:
        y_end = min(y + split_height, height)
        result.append((y, y_end))
        if y_end == height:
            break
        y += step
    return result


def split_one(image_path: Path, out_dir: Path, options: SplitOptions) -> int:
    with Image.open(image_path) as img:
        src = img.convert("RGB") if options.output_format in {"jpg", "jpeg"} else img.copy()
        width, height = src.size

        if options.mode == "cuts":
            segments = split_by_cuts(height, options.cuts)
        else:
            segments = split_by_height(height, options.split_height, options.overlap)

        ext = "jpg" if options.output_format == "jpeg" else options.output_format
        count = 0
        for idx, (y0, y1) in enumerate(segments, start=1):
            out_file = out_dir / f"{image_path.stem}_part{idx:03d}_{y0}-{y1}.{ext}"
            cropped = src.crop((0, y0, width, y1))
            save_kwargs = {"quality": options.jpeg_quality, "optimize": True} if ext == "jpg" else {}
            cropped.save(out_file, **save_kwargs)
            count += 1
        return count


class DesktopTool:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("画像分割カットツール（実運用版）")
        self.root.geometry("1050x760")

        self.inputs: list[Path] = []
        self.settings = self.load_settings()

        self.mode_var = StringVar(value="cuts")
        self.cuts_var = StringVar(value="1000,2000,3000")
        self.height_var = StringVar(value="1200")
        self.overlap_var = StringVar(value="0")
        self.format_var = StringVar(value="jpg")
        self.quality_var = StringVar(value="92")
        self.workers_var = StringVar(value="4")
        self.provider_var = StringVar(value="auto")
        self.target_var = StringVar(value="JA")
        self.deepl_key_var = StringVar(value=self.settings.get("deepl_api_key", ""))
        self.gemini_key_var = StringVar(value=self.settings.get("gemini_api_key", ""))

        default_output = APP_DIR / "output" / datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_var = StringVar(value=str(default_output))

        self._build_ui()

    def load_settings(self) -> dict:
        if SETTINGS_PATH.exists():
            try:
                return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_settings(self) -> None:
        payload = {
            "deepl_api_key": self.deepl_key_var.get().strip(),
            "gemini_api_key": self.gemini_key_var.get().strip(),
        }
        SETTINGS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log("APIキーを settings.json に保存しました。")

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=True)

        paned = ttk.Panedwindow(main, orient="horizontal")
        paned.pack(fill=BOTH, expand=True)

        left = ttk.Frame(paned, padding=8)
        right = ttk.Frame(paned, padding=8)
        paned.add(left, weight=2)
        paned.add(right, weight=1)

        self.build_split_frame(left)
        self.build_translate_frame(right)

    def build_split_frame(self, parent: ttk.Frame) -> None:
        files_frame = ttk.LabelFrame(parent, text="入力画像", padding=8)
        files_frame.pack(fill=BOTH, expand=True)

        btn_row = ttk.Frame(files_frame)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="画像追加", command=self.add_files).pack(side=LEFT)
        ttk.Button(btn_row, text="フォルダ追加", command=self.add_folder).pack(side=LEFT, padx=6)
        ttk.Button(btn_row, text="クリア", command=self.clear_inputs).pack(side=LEFT)

        self.file_list = ttk.Treeview(files_frame, columns=("path",), show="headings", height=10)
        self.file_list.heading("path", text="ファイル")
        self.file_list.column("path", width=620)
        self.file_list.pack(fill=BOTH, expand=True, pady=8)

        opt = ttk.LabelFrame(parent, text="分割設定", padding=8)
        opt.pack(fill="x", pady=8)

        ttk.Radiobutton(opt, text="カット座標", value="cuts", variable=self.mode_var).grid(row=0, column=0, sticky="w")
        ttk.Entry(opt, textvariable=self.cuts_var, width=40).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="例: 1200,2400,3600").grid(row=0, column=2, sticky="w")

        ttk.Radiobutton(opt, text="固定高さ", value="height", variable=self.mode_var).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(opt, textvariable=self.height_var, width=12).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="overlap").grid(row=1, column=2, sticky="e")
        ttk.Entry(opt, textvariable=self.overlap_var, width=8).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(opt, text="format").grid(row=2, column=0, sticky="w")
        ttk.Combobox(opt, values=["jpg", "png", "webp"], textvariable=self.format_var, width=10, state="readonly").grid(row=2, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="jpeg quality").grid(row=2, column=2, sticky="e")
        ttk.Entry(opt, textvariable=self.quality_var, width=8).grid(row=2, column=3, sticky="w", padx=4)
        ttk.Label(opt, text="workers").grid(row=2, column=4, sticky="e")
        ttk.Entry(opt, textvariable=self.workers_var, width=8).grid(row=2, column=5, sticky="w", padx=4)

        out = ttk.LabelFrame(parent, text="出力", padding=8)
        out.pack(fill="x")
        ttk.Entry(out, textvariable=self.output_var).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(out, text="参照", command=self.pick_output).pack(side=LEFT, padx=6)

        ttk.Button(parent, text="✂ 分割開始", command=self.run_split, style="Accent.TButton").pack(fill="x", pady=10)

        log_box = ttk.LabelFrame(parent, text="ログ", padding=6)
        log_box.pack(fill=BOTH, expand=True)
        self.log_text = ScrolledText(log_box, height=10)
        self.log_text.pack(fill=BOTH, expand=True)

    def build_translate_frame(self, parent: ttk.Frame) -> None:
        api = ttk.LabelFrame(parent, text="翻訳API設定", padding=8)
        api.pack(fill="x")
        ttk.Label(api, text="DeepL API Key").grid(row=0, column=0, sticky="w")
        ttk.Entry(api, textvariable=self.deepl_key_var, show="*", width=40).grid(row=1, column=0, sticky="we", pady=4)
        ttk.Label(api, text="Gemini API Key").grid(row=2, column=0, sticky="w")
        ttk.Entry(api, textvariable=self.gemini_key_var, show="*", width=40).grid(row=3, column=0, sticky="we", pady=4)
        ttk.Button(api, text="キー保存", command=self.save_settings).grid(row=4, column=0, sticky="w", pady=(4, 0))

        tr = ttk.LabelFrame(parent, text="翻訳（実運用）", padding=8)
        tr.pack(fill=BOTH, expand=True, pady=8)

        top_row = ttk.Frame(tr)
        top_row.pack(fill="x")
        ttk.Label(top_row, text="Provider").pack(side=LEFT)
        ttk.Combobox(top_row, values=["auto", "deepl", "gemini"], textvariable=self.provider_var, state="readonly", width=10).pack(side=LEFT, padx=4)
        ttk.Label(top_row, text="Target").pack(side=LEFT, padx=(10, 0))
        ttk.Combobox(top_row, values=["JA", "EN", "ZH", "KO"], textvariable=self.target_var, state="readonly", width=8).pack(side=LEFT, padx=4)

        self.src_text = ScrolledText(tr, height=9)
        self.src_text.pack(fill=BOTH, expand=True, pady=6)
        self.dst_text = ScrolledText(tr, height=9)
        self.dst_text.pack(fill=BOTH, expand=True)

        ttk.Button(tr, text="翻訳実行", command=self.run_translate).pack(fill="x", pady=6)

    def log(self, text: str) -> None:
        self.log_text.insert(END, f"{datetime.now():%H:%M:%S} {text}\n")
        self.log_text.see(END)

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff")])
        for raw in selected:
            path = Path(raw)
            if path not in self.inputs:
                self.inputs.append(path)
                self.file_list.insert("", END, values=(str(path),))

    def add_folder(self) -> None:
        folder = filedialog.askdirectory()
        if not folder:
            return
        for ext in ("*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp", "*.tif", "*.tiff"):
            for f in sorted(Path(folder).glob(ext)):
                if f not in self.inputs:
                    self.inputs.append(f)
                    self.file_list.insert("", END, values=(str(f),))

    def clear_inputs(self) -> None:
        self.inputs.clear()
        for iid in self.file_list.get_children():
            self.file_list.delete(iid)

    def pick_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def collect_options(self) -> SplitOptions:
        mode = self.mode_var.get()
        cuts = parse_cuts(self.cuts_var.get()) if mode == "cuts" else []
        split_height = int(self.height_var.get())
        overlap = int(self.overlap_var.get())
        out_format = self.format_var.get()
        quality = int(self.quality_var.get())
        workers = max(1, int(self.workers_var.get()))

        if not 1 <= quality <= 100:
            raise ValueError("jpeg quality は 1-100")

        return SplitOptions(
            mode=mode,
            cuts=cuts,
            split_height=split_height,
            overlap=overlap,
            output_format=out_format,
            jpeg_quality=quality,
            workers=workers,
        )

    def run_split(self) -> None:
        if not self.inputs:
            messagebox.showwarning("入力不足", "画像を追加してください")
            return

        try:
            options = self.collect_options()
        except Exception as e:
            messagebox.showerror("設定エラー", str(e))
            return

        out_dir = Path(self.output_var.get()).expanduser()
        out_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"分割開始: {len(self.inputs)}枚 / 出力: {out_dir}")

        def worker() -> None:
            try:
                total = 0
                with ThreadPoolExecutor(max_workers=options.workers) as ex:
                    futures = [ex.submit(split_one, p, out_dir, options) for p in self.inputs]
                    for p, fut in zip(self.inputs, futures):
                        count = fut.result()
                        total += count
                        self.root.after(0, self.log, f"完了: {p.name} -> {count}枚")
                self.root.after(0, self.log, f"全完了: 合計 {total} ファイル")
                self.root.after(0, lambda: messagebox.showinfo("完了", f"分割完了: {total} ファイル"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("処理エラー", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def translate_deepl(self, text: str, target: str) -> str:
        key = self.deepl_key_var.get().strip()
        if not key:
            raise ValueError("DeepL APIキーが未設定")

        endpoint = "https://api-free.deepl.com/v2/translate" if key.endswith(":fx") else "https://api.deepl.com/v2/translate"
        resp = requests.post(
            endpoint,
            headers={"Authorization": f"DeepL-Auth-Key {key}"},
            data={"text": text, "target_lang": target},
            timeout=40,
        )
        resp.raise_for_status()
        data = resp.json()
        translated = (data.get("translations") or [{}])[0].get("text")
        if not translated:
            raise RuntimeError("DeepL翻訳結果が空")
        return translated

    def translate_gemini(self, text: str, target: str) -> str:
        key = self.gemini_key_var.get().strip()
        if not key:
            raise ValueError("Gemini APIキーが未設定")

        target_map = {"JA": "日本語", "EN": "英語", "ZH": "中国語", "KO": "韓国語"}
        lang = target_map.get(target, target)
        prompt = f"次の文章を{lang}に自然に翻訳し、翻訳文のみ返してください。\n\n{text}"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={key}"
        resp = requests.post(
            url,
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.2},
            },
            timeout=40,
        )
        resp.raise_for_status()
        data = resp.json()
        translated = (((data.get("candidates") or [{}])[0].get("content") or {}).get("parts") or [{}])[0].get("text")
        if not translated:
            raise RuntimeError("Gemini翻訳結果が空")
        return translated.strip()

    def run_translate(self) -> None:
        text = self.src_text.get("1.0", END).strip()
        if not text:
            messagebox.showwarning("入力不足", "翻訳テキストを入力してください")
            return

        provider = self.provider_var.get()
        target = self.target_var.get()

        try:
            if provider == "deepl":
                translated = self.translate_deepl(text, target)
            elif provider == "gemini":
                translated = self.translate_gemini(text, target)
            else:
                try:
                    translated = self.translate_deepl(text, target)
                except Exception:
                    translated = self.translate_gemini(text, target)
            self.dst_text.delete("1.0", END)
            self.dst_text.insert("1.0", translated)
        except Exception as e:
            messagebox.showerror("翻訳エラー", str(e))


def main() -> None:
    root = Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = DesktopTool(root)
    app.log("起動完了")
    root.mainloop()


if __name__ == "__main__":
    main()
