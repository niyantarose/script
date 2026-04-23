# 画像分割カットツール（desktop_tool / 実運用版）

このフォルダをそのまま置けば、Windowsで実行できます。  
GASの制限なしで、**画像分割 + 翻訳(DeepL / Gemini)** を1つのツールで運用できます。

## フォルダ構成

- `app.py` : デスクトップGUI本体（Tkinter）
- `requirements.txt` : 依存ライブラリ
- `run.bat` : 初回セットアップ込み起動バッチ
- `settings.json` : APIキー保存先（初回実行後に自動作成）

## 使い方（Windows）

1. `run.bat` をダブルクリック
2. 左側で画像を追加（複数可 / フォルダ追加可）
3. 分割モードを選択
   - カット座標: `1200,2400,3600`
   - 固定高さ: `1200` + `overlap`
4. 出力フォルダを指定して「分割開始」

## 翻訳機能

- DeepL APIキー / Gemini APIキーを入力して「キー保存」
- Provider は `auto / deepl / gemini`
- `auto` は **DeepL優先 → 失敗時Gemini**

## 実運用のおすすめ

- 翻訳品質優先: `deepl`
- 停止回避優先: `auto`
- 処理速度優先: workers を 4〜8 で調整

## 注意

- APIキーは `settings.json` に保存されます。共有PCでは管理に注意してください。
- DeepL Freeキー（末尾 `:fx`）は Freeエンドポイントへ自動切替します。
