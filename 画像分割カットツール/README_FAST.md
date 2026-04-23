# 画像分割カットツール（ローカル高速版）

GAS(Webアプリ)の実行制限・タイムアウトを避けるための、**ローカル実行版**です。  
`Pillow` で直接画像を処理するため、GASより高速に大量分割できます。

## 1) セットアップ

```bash
cd "C:\Users\Owner\Desktop\script\画像分割カットツール"
py -m venv .venv
.venv\Scripts\activate
pip install pillow
```

## 2) 使い方

### A. Y座標を指定して分割（手動線と同じイメージ）

```bash
python fast_splitter.py "入力画像フォルダ" --cuts 1200,2400,3600 -o output
```

### B. 高さ固定で自動分割（例: 1500pxごと）

```bash
python fast_splitter.py "入力画像フォルダ" --split-height 1500 --overlap 0 -o output
```

## 3) よく使うオプション

- `--format png|jpg|webp` 出力形式
- `--jpeg-quality 92` JPEG品質
- `--workers 8` 並列数（CPUに合わせて調整）

例:

```bash
python fast_splitter.py ./input --split-height 1500 --format jpg --jpeg-quality 90 --workers 8 -o ./output
```

## 4) 出力ファイル名

`元ファイル名_part001_0-1500.png` のように、分割番号とY範囲を付けて保存します。

---

必要なら次の段階で、いまの `index.html` UI と同じ見た目の **ローカルデスクトップ版(Electron / Python GUI)** にもできます。
