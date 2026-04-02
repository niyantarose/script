# にゃんたろうず商品登録拡張機能

`Z:\script\台湾拡張機能` と `Z:\script\アラジン拡張機能\アラジンウェブアプリ本体` を 1 本の拡張へ統合した作業版です。

## 現在の実装

- popup ルーター
  - `popup/popup.html` が現在タブを判定し、台湾またはアラジンの既存 popup へ転送
- 台湾機能
  - `popup/taiwan/*`
  - `backgrounds/taiwan.js`
  - `content/taiwan/content.js`
- アラジン機能
  - `popup/aladin/*`
  - `backgrounds/aladin.js`
  - `content/aladin/content.js`

## 統合で行ったこと

- manifest を 1 本化
- service worker から台湾・アラジンの background を両方読み込む構成へ変更
- `downloadImage` / `downloadCsv` の generic action は 1 系統だけが受けるように調整
- content 注入パスを統合拡張内の場所へ合わせて修正
- popup は現在ページで台湾 / アラジンを自動判定

## 参照元

- `Z:\script\台湾拡張機能`
- `Z:\script\アラジン拡張機能\アラジンウェブアプリ本体`
- `Z:\script\韓国グッズ拡張機能`

## まだ残っているもの

- `core/`, `genres/`, `normalizers/` などの新アーキテクチャ用 scaffold
  - 今回は削除せず残しています
  - 実動作は現在 `popup/*`, `backgrounds/*`, `content/*` の移植版が中心です

## 次に見る場所

1. `manifest.json`
2. `service-worker.js`
3. `backgrounds/taiwan.js`
4. `backgrounds/aladin.js`
5. `popup/taiwan/popup.html`
6. `popup/aladin/popup.html`

## ダウンロード競合回避

- onDeterminingFilename のような global hook は使わない
- 保存時は各要求ごとに chrome.downloads.download() を呼ぶだけにする
- これにより Imageye / Imageya のような他のダウンロード系拡張と競合しにくい構成を維持する
