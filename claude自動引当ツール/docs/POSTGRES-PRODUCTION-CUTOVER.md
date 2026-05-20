# 本番 PostgreSQL 切替手順書（実行前の整理用）

**この文書は手順の整理です。メンテナンス窓で実施し、各コマンド実行前に担当者でダブルチェックしてください。**  
本番 DB 名・接続先は **毎回表示を確認してから** 実行すること（`zaiko_inventory` と `zaiko_inventory_test` の取り違え防止）。

---

## 対象環境（想定）

| 項目 | 値 |
|------|-----|
| アプリルート | `/home/ubuntu/zaiko-tool/app` |
| SQLite（現本番） | `/home/ubuntu/zaiko-tool/app/instance/inventory.db` |
| systemd（Gunicorn） | `zaiko-tool.service`（system） |
| 定期ジョブ | `systemctl --user` の Yahoo 取込・在庫系 timer / oneshot |
| PostgreSQL 本番 DB 名（想定） | `zaiko_inventory` |
| PostgreSQL テスト DB 名（参考） | `zaiko_inventory_test` |
| DB ロール（想定） | `zaiko_app` |

---

## 1. 本番切替前チェックリスト

作業開始前に、次をすべて確認してからチェックを付ける。

- [ ] **Linger**  
  `loginctl show-user ubuntu -p Linger` が `Linger=yes`（user timer がログアウト後も動く前提）。
- [ ] **systemd timer 状態**  
  `systemctl --user list-timers` で対象 timer の `NEXT` / `LAST` が想定どおり。直近の失敗がないこと（`systemctl --user list-units 'zaiko-*'` 等）。
- [ ] **PostgreSQL 稼働**  
  `sudo systemctl status postgresql`（またはディストリビューションに合うユニット名）が active。
- [ ] **`zaiko_app` 権限**  
  本番 DB `zaiko_inventory` に対し `CONNECT` / `CREATE`（初回のみ管理者）/ `USAGE` on `public` / テーブル操作が可能なこと（方針に合わせ事前検証）。
- [ ] **空の本番 DB `zaiko_inventory` 作成可否**  
  既に存在する場合は **中身がテスト用でないか**、**本番切替用の空 DB か** を確認。**誤った DB を `DROP` しないこと。**
- [ ] **`.env` バックアップ**  
  例: `cp -a /home/ubuntu/zaiko-tool/app/.env /home/ubuntu/zaiko-tool/app/.env.bak.$(date +%Y%m%d%H%M)`  
  復旧に必要なキー（`SECRET_KEY`、Yahoo 系、現行 `USE_SQLITE`）が含まれることを確認。
- [ ] **SQLite 最終バックアップ**  
  メンテ開始直前に **アプリ停止後** または **書き込みが止まったタイミング** でコピー（下記手順の「SQLite 最終バックアップ」参照）。
- [ ] **ロールバック用 SQLite 保存場所**  
  バックアップファイルを **本番サーバ上の専用ディレクトリ** に保存（例: `/home/ubuntu/zaiko-tool/backups/`）。ファイル名に日時を含める。  
  切替成功後も **一定期間は削除しない**（ロールバック・監査用）。

---

## 2. 本番切替手順（概要）

1. user **timer 停止**（取込・在庫ジョブが走らない状態にする）。
2. **`zaiko-tool.service` 停止**（Gunicorn / SQLite ロック解放）。
3. **SQLite 最終バックアップ**（この時点のファイルを「本番ロールバックの正」とする）。
4. **本番 DB `zaiko_inventory` の用意**（空 DB 新規作成、または空の `public` のみ）。
5. **スキーマ初期化**（アプリの `create_all` 相当。既存テーブルがある場合は方針要確認）。
6. **移行スクリプト**で SQLite 最終バックアップ → PostgreSQL 本番へ投入（`--precheck-only` → 本移行）。
7. **`.env` 切替**（`USE_SQLITE=false`、`DATABASE_URL=postgresql+psycopg://...` 本番向け）。
8. **`zaiko-tool.service` 再起動**。
9. **画面・API の確認**（主要 URL、手動ジョブ）。
10. **user timer 再開**。

---

## 3. ロールバック手順（概要）

1. user **timer 停止**。
2. **`zaiko-tool.service` 停止**。
3. **`.env` を SQLite 運用に戻す**（`USE_SQLITE=true`、SQLite 用 URI。バックアップした `.env.bak.*` を参照）。
4. **`inventory.db` を切替前バックアップに戻す**（ファイルコピーで上書き。**戻すファイルが正しいバックアップか必ず確認**）。
5. **`zaiko-tool.service` 再起動**。
6. **簡易動作確認**（トップ・受注一覧など）。
7. **user timer 再開**。

PostgreSQL 側のデータは **即削除しない**（調査・再試行用に温存可）。アプリは SQLite を向ければ業務は継続できる。

---

## 4. 本番切替時に使う具体コマンド（実行順・パスワードは伏せる）

**以下の `***` はパスワード等の伏字です。実際の値は `.env` または秘密管理に従うこと。**  
**`DATABASE_URL` や `psql -d` のデータベース名は、必ず端末表示で `zaiko_inventory` であることを確認してから実行すること。**

### 4.1 事前：接続・DB 名の確認（本番 DB 名確認）

```text
1. sudo -u postgres psql -c '\l'
   → 本番用は zaiko_inventory、テスト用は zaiko_inventory_test など、名前を取り違えないこと。
```

### 4.2 user timer 停止

```text
2. systemctl --user stop zaiko-yahoo-orders.timer
3. systemctl --user stop zaiko-yahoo-stock-diff.timer
4. systemctl --user stop zaiko-yahoo-stock-full.timer
5. systemctl --user stop zaiko-run-all-imports.timer
   （実際に使っている timer 名に合わせて追加・削除）
```

### 4.3 Gunicorn（本番サービス）停止

```text
6. sudo systemctl stop zaiko-tool.service
7. sudo systemctl status zaiko-tool.service --no-pager
   → inactive であること
```

### 4.4 SQLite 最終バックアップ

```text
8. install -d -m 700 /home/ubuntu/zaiko-tool/backups
9. TS=$(date +%Y%m%d%H%M)
10. cp -a /home/ubuntu/zaiko-tool/app/instance/inventory.db \
      /home/ubuntu/zaiko-tool/backups/inventory.pre-pg-cutover.${TS}.db
11. ls -l /home/ubuntu/zaiko-tool/backups/inventory.pre-pg-cutover.${TS}.db
    → サイズ・更新時刻が想定どおりか確認
```

以降の移行では **`DB_COPY=/home/ubuntu/zaiko-tool/backups/inventory.pre-pg-cutover.${TS}.db`** を読み取り専用の移行元とする。

### 4.5 本番 DB 空作成（未作成の場合のみ・本番 DB 名確認）

**危険度：中（誤った DB 名で実行しないこと）**

```text
12. sudo -u postgres psql -c "CREATE DATABASE zaiko_inventory OWNER zaiko_app;"
    → 既に存在する場合はスキップ。CREATE の前に \l で名前を再確認。
```

### 4.6 本番 DB の public スキーマ初期化（空にする・本番 DB 名確認）

**危険度：高（DROP SCHEMA は対象 DB が zaiko_inventory であることを psql のプロンプトや `-d` で必ず確認）**

```text
13. sudo -u postgres psql -d zaiko_inventory <<'SQL'
    DROP SCHEMA public CASCADE;
    CREATE SCHEMA public;
    GRANT ALL ON SCHEMA public TO zaiko_app;
    GRANT ALL ON SCHEMA public TO public;
    SQL
```

既に本番データが入っている DB に対して実行すると **全テーブルが削除**される。**テスト DB ではないこと**を再確認。

### 4.7 移行スクリプト（本番 `.env.pgprod` 等で DATABASE_URL を本番に向ける）

**危険度：中（--env-file の DATABASE_URL が本番 zaiko_inventory を指すこと）**

事前に **本番用** の環境ファイルを用意する（例: `/home/ubuntu/zaiko-tool/app/.env.pgprod`）。  
内容例（パスワードは伏せる）:

```env
USE_SQLITE=false
DATABASE_URL=postgresql+psycopg://zaiko_app:***@127.0.0.1:5432/zaiko_inventory
SECRET_KEY=（既存 .env と同じ）
（Yahoo 等、既存 .env と同じキーを必要分コピー）
```

```text
14. cd /home/ubuntu/zaiko-tool/app
15. source venv/bin/activate
16. export DB_COPY=/home/ubuntu/zaiko-tool/backups/inventory.pre-pg-cutover.${TS}.db
17. python scripts/migrate_sqlite_to_postgres.py \
      --sqlite-path "$DB_COPY" \
      --env-file /home/ubuntu/zaiko-tool/app/.env.pgprod \
      --precheck-only
18. python scripts/migrate_sqlite_to_postgres.py \
      --sqlite-path "$DB_COPY" \
      --env-file /home/ubuntu/zaiko-tool/app/.env.pgprod
    → 件数一致・exit 0 を確認
```

### 4.8 本番 `.env` 切替

```text
19. cp -a /home/ubuntu/zaiko-tool/app/.env /home/ubuntu/zaiko-tool/app/.env.sqlite-final.bak.${TS}
20. （エディタで）/home/ubuntu/zaiko-tool/app/.env を編集:
     USE_SQLITE=false
     DATABASE_URL=postgresql+psycopg://zaiko_app:***@127.0.0.1:5432/zaiko_inventory
    または DB_HOST / DB_USER / DB_PASSWORD / DB_NAME で接続（config 方針に合わせる）
21. grep -E '^USE_SQLITE|^DATABASE_URL|^DB_' /home/ubuntu/zaiko-tool/app/.env
    → 意図した値のみであること
```

### 4.9 Gunicorn 再起動

```text
22. sudo systemctl daemon-reload
    （unit ファイルを変えた場合のみ必須）
23. sudo systemctl start zaiko-tool.service
24. sudo systemctl status zaiko-tool.service --no-pager
25. sudo journalctl -u zaiko-tool.service -n 80 --no-pager
```

### 4.10 画面・手動ジョブ確認後、timer 再開

```text
26. （ブラウザで主要 URL を確認。詳細は「5. 本番切替後の確認項目」）
27. systemctl --user start zaiko-yahoo-orders.timer
28. systemctl --user start zaiko-yahoo-stock-diff.timer
29. systemctl --user start zaiko-yahoo-stock-full.timer
30. systemctl --user start zaiko-run-all-imports.timer
31. systemctl --user list-timers --no-pager
```

---

## 5. 本番切替後の確認項目

### 5.1 ブラウザ（想定パス）

- [ ] `/dashboard`（またはアプリのトップ）
- [ ] `/orders/`
- [ ] `/order-search/`
- [ ] `/stock`
- [ ] `/japan/`
- [ ] `/ems/?agent=daniel`
- [ ] `/purchases/?agent=daniel`

### 5.2 ログ・プロセス

- [ ] `sudo journalctl -u zaiko-tool.service --since "10 minutes ago" --no-pager`  
  致命的エラー・大量の DB エラーがないこと。
- [ ] `journalctl --user -u zaiko-yahoo-orders.service -n 50 --no-pager`（他 timer も必要に応じて）
- [ ] `systemctl --user list-timers` で `NEXT` が妥当であること。

### 5.3 PostgreSQL 件数（本番 DB 名確認）

```text
sudo -u postgres psql -d zaiko_inventory -c "SELECT 'orders' AS t, COUNT(*) FROM orders UNION ALL SELECT 'order_items', COUNT(*) FROM order_items UNION ALL SELECT 'inventory', COUNT(*) FROM inventory;"
```

SQLite 最終バックアップと **主要テーブル件数が整合**していること（移行スクリプト出力と突合）。

### 5.4 手動ジョブ（運用方針に従い）

- [ ] Yahoo 注文取込（HTTP または運用中の方法）
- [ ] 在庫差分（HTTP）
- [ ] 在庫フル（CLI / user oneshot）

---

## 6. ロールバック用コマンド（番号は独立した手順として実行）

**本番 DB 名に触れず、SQLite に戻すだけの流れ。**

```text
R1. systemctl --user stop zaiko-yahoo-orders.timer zaiko-yahoo-stock-diff.timer zaiko-yahoo-stock-full.timer zaiko-run-all-imports.timer
R2. sudo systemctl stop zaiko-tool.service
R3. cp -a /home/ubuntu/zaiko-tool/app/.env /home/ubuntu/zaiko-tool/app/.env.pg-failed.bak.$(date +%Y%m%d%H%M)
R4. cp -a /home/ubuntu/zaiko-tool/app/.env.sqlite-final.bak.* /home/ubuntu/zaiko-tool/app/.env
    （切替直前に保存した SQLite 用 .env の実ファイル名に合わせる）
R5. cp -a /home/ubuntu/zaiko-tool/backups/inventory.pre-pg-cutover.${TS}.db \
      /home/ubuntu/zaiko-tool/app/instance/inventory.db
    → ${TS} は切替前に取ったバックアップの日時に合わせる
R6. sudo systemctl start zaiko-tool.service
R7. sudo systemctl status zaiko-tool.service --no-pager
R8. （トップ・受注一覧を確認）
R9. systemctl --user start （各 timer）
```

---

## 7. まだやらないこと（本手順の範囲外）

- Alembic / Flask-Migrate の本格導入
- Redis 導入
- スキーマの大規模改修（本手順は現行モデル前提の移行）
- **古い SQLite の削除**（バックアップは保持）
- PostgreSQL の **外部公開**（`listen_addresses` やファイアウォールは引き続き localhost 前提）

---

## 8. 参考（既存ドキュメント）

- テスト移行・スキーマ初期化の考え方: `docs/POSTGRES-MIGRATION-TEST.md`
- 移行スクリプト: `scripts/migrate_sqlite_to_postgres.py`
