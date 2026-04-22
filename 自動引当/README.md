# 自動引当ツール

`要件定義書_v2.docx`、`DB設計書_v2.docx`、`環境構築手順書.docx` を参照しながら進めている、Flask + SQLAlchemy ベースの在庫自動引当ツールです。

現時点で入っている土台:

- ダッシュボード
- ダニエル発注リスト
- ダニエルEMSリスト
- テグ発注リスト
- テグEMSリスト
- Yahoo受注リスト
- 日本在庫管理
- 注文検索
- imported_files による取込ファイル重複防止の土台
- Ubuntu / MySQL / Nginx / systemd 用の配置ファイル

## 参照した設計書の要点

- DB設計書: 9テーブル構成を基準に、追加要件として `imported_files` を実装
- 要件定義書: 2段階引当、4段階漏れチェック、遅延管理、日本在庫反映の流れを反映
- 環境構築手順書: Ubuntu 22.04 + MySQL + Flask + Nginx + systemd の配置を前提化

補足:

- ダニエル / テグのページ分離を実現するため、`purchases.source_type` と `ems.source_type` を追加しています。これは画面要件に合わせた実装上の補助列です。
- 以前のご要望に合わせて、画面上の「商品名」はまだ出していません。

## ローカル起動

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
Copy-Item .env.example .env
.\.venv\Scripts\python app.py
```

ブラウザ:

- [http://127.0.0.1:5000](http://127.0.0.1:5000)

## 便利コマンド

```powershell
$env:FLASK_APP='app.py'
.\.venv\Scripts\flask seed-demo
.\.venv\Scripts\flask recalculate
.\.venv\Scripts\flask run-checks
.\.venv\Scripts\flask import-all
.\.venv\Scripts\flask import-kind daniel_purchases
```

## 本番配置ファイル

- Gunicorn設定: [gunicorn.conf.py](/C:/Users/Owner/Desktop/script/自動引当/gunicorn.conf.py)
- Ubuntuセットアップ: [deploy/ubuntu/setup_vps.sh](/C:/Users/Owner/Desktop/script/自動引当/deploy/ubuntu/setup_vps.sh)
- systemd: [deploy/systemd/inventory-tool.service](/C:/Users/Owner/Desktop/script/自動引当/deploy/systemd/inventory-tool.service)
- Nginx: [deploy/nginx/inventory-tool.conf](/C:/Users/Owner/Desktop/script/自動引当/deploy/nginx/inventory-tool.conf)
- cron: [deploy/cron/hourly_import.cron](/C:/Users/Owner/Desktop/script/自動引当/deploy/cron/hourly_import.cron)

## いまの取込の状態

外部APIの実接続そのものはまだ未実装です。現状の `データ取込` ボタンと CLI は、設計書どおりのデータ源を置くための土台として、以下を先に持っています。

- 取込種別ごとのルーティング
- 取込済みファイル名の記録
- 同一ファイル名の重複取込防止
- ダニエル / テグ / Yahoo の画面別導線

次に実装する対象:

1. Yahoo API 実接続
2. Cloudike WebDAV 実接続
3. Google Sheets API 実接続
4. cron の本番設定
5. インライン編集と編集ログ
