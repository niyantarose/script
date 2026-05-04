# PostgreSQL テスト移行（本番切替なし）

本番 SQLite（`instance/inventory.db`）は**書き換えず**、テスト用 PostgreSQL にコピーして検証する手順です。

## 前提

- VPS 内の PostgreSQL のみ利用し、`listen_addresses` は localhost 相当に限定する。
- 本番 `.env` の `USE_SQLITE=true` は**このドキュメントの作業では変更しない**。
- アプリコードは `scp` で VPS の `/home/ubuntu/zaiko-tool/app` へ反映する想定。

## 1. PostgreSQL のインストール（VPS / Ubuntu 例）

```bash
sudo apt update
sudo apt install -y postgresql postgresql-contrib
sudo systemctl enable --now postgresql
```

## 2. テスト DB とロールの作成（例）

```bash
sudo -u postgres psql -c "CREATE USER zaiko_app WITH PASSWORD '（強いパスワード）';"
sudo -u postgres psql -c "CREATE DATABASE zaiko_inventory_test OWNER zaiko_app;"
sudo -u postgres psql -d zaiko_inventory_test -c "GRANT ALL ON SCHEMA public TO zaiko_app;"
```

## 3. venv に依存追加

プロジェクトルート（または VPS の `app` ディレクトリで venv を有効化したうえで）:

```bash
pip install 'psycopg[binary]>=3.2.0'
```

または `requirements.txt` を更新済みなら:

```bash
pip install -r requirements.txt
```

## 4. SQLite のバックアップ（読み取り専用コピー推奨）

本番ファイルを触らずコピーして移行元にする:

```bash
cp -a /home/ubuntu/zaiko-tool/app/instance/inventory.db \
  /home/ubuntu/zaiko-tool/app/instance/inventory.precopy.$(date +%Y%m%d%H%M).db
```

移行スクリプトには**このコピー**のパスを `--sqlite-path` で渡すとより安全です。

## 5. `.env.pgtest` の用意

```bash
cd /home/ubuntu/zaiko-tool/app
cp .env.pgtest.example .env.pgtest
# エディタで DATABASE_URL と SECRET_KEY 等を編集
```

## 6. 事前チェック（SQLite 読み取りのみ）

```bash
cd /home/ubuntu/zaiko-tool/app
source venv/bin/activate
python scripts/migrate_sqlite_to_postgres.py \
  --sqlite-path /home/ubuntu/zaiko-tool/app/instance/inventory.precopy.YYYYMMDD.db \
  --env-file /home/ubuntu/zaiko-tool/app/.env.pgtest \
  --precheck-only
```

重複・孤立 FK があれば exit code 1。修正するか、移行を見送る判断をする。

## 7. テスト移行の実行

対象 PG が空であること（既にデータがある場合は `--truncate` で**テスト DBのみ**全消去してから投入）:

```bash
python scripts/migrate_sqlite_to_postgres.py \
  --sqlite-path /path/to/inventory.precopy.db \
  --env-file /home/ubuntu/zaiko-tool/app/.env.pgtest
```

成功するとテーブルごとの件数比較と `setval` 補正が表示され、exit code 0。

## 8. アプリのテスト起動（任意）

別シェルで一時的に PG を向けたい場合のみ（**本番 systemd は触らない**）:

```bash
export $(grep -v '^#' .env.pgtest | xargs)
# USE_SQLITE=false DATABASE_URL=... が有効な状態で
gunicorn -w 1 --bind 127.0.0.1:5001 app:app
```

ブラウザで `http://VPS:5001` など（ファイアウォール注意）で画面確認。

## 9. ロールバック方針

- 本番 SQLite は変更していないため、**アプリを SQLite に戻すだけ**で従来運用に戻る。
- テスト用 PostgreSQL のデータは削除しなくてよい（再検証用に温存可）。
- `DROP DATABASE zaiko_inventory_test;` は**検証完了後に任意**（本番 DB 名と取り違えないこと）。

## 10. 今回やらないこと

- 本番 `.env` の `USE_SQLITE` 切替
- 本番 PostgreSQL への本番 DB 作成・切替
- SQLite ファイルの削除
- user systemd timer の変更
- Alembic の本格運用開始

## 11. Windows から VPS へ `scp` するファイル一覧（例）

リポジトリルート: `C:\Users\Owner\Desktop\script\claude自動引当ツール\`  
VPS アプリルート: `/home/ubuntu/zaiko-tool/app/`  
（`USER@VPS` は接続先に置き換え）

```powershell
scp -i "C:\Users\Owner\.ssh\id_ed25519" `
  "C:\Users\Owner\Desktop\script\claude自動引当ツール\config.py" `
  ubuntu@VPS:/home/ubuntu/zaiko-tool/app/config.py

scp -i "C:\Users\Owner\.ssh\id_ed25519" `
  "C:\Users\Owner\Desktop\script\claude自動引当ツール\requirements.txt" `
  ubuntu@VPS:/home/ubuntu/zaiko-tool/app/requirements.txt

scp -i "C:\Users\Owner\.ssh\id_ed25519" `
  "C:\Users\Owner\Desktop\script\claude自動引当ツール\scripts\migrate_sqlite_to_postgres.py" `
  ubuntu@VPS:/home/ubuntu/zaiko-tool/app/scripts/migrate_sqlite_to_postgres.py
```

ドキュメント・サンプル（任意）:

```powershell
scp -i "C:\Users\Owner\.ssh\id_ed25519" `
  "C:\Users\Owner\Desktop\script\claude自動引当ツール\.env.pgtest.example" `
  ubuntu@VPS:/home/ubuntu/zaiko-tool/app/.env.pgtest.example

scp -i "C:\Users\Owner\.ssh\id_ed25519" `
  "C:\Users\Owner\Desktop\script\claude自動引当ツール\docs\POSTGRES-MIGRATION-TEST.md" `
  ubuntu@VPS:/home/ubuntu/zaiko-tool/app/docs/POSTGRES-MIGRATION-TEST.md
```

VPS で `docs` が無ければ `mkdir -p ~/zaiko-tool/app/docs` を先に実行。
