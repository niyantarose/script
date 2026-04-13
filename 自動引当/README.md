# 自動引当ツール

要件定義書・DB設計書をもとに作成した、Flask + SQLAlchemy ベースの自動引当ツールです。  
このフォルダ単体で起動でき、以下を一通り試せます。

- 7画面のWeb UI
- 即納優先の自動引当再計算
- 韓国発注の仮引当
- EMS登録時の本引当
- 4段階漏れチェックと遅延アラート
- 日本在庫仕分けと即納在庫への反映
- 画面単位のCSV出力

## フォルダ構成

```text
app.py
inventory_tool/
templates/
static/
create_tables.sql
requirements.txt
```

## 初回セットアップ

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

`.env` の `DATABASE_URL` はデフォルトで SQLite を使います。  
本番で MySQL を使う場合は、例えば次のように変更します。

```env
DATABASE_URL=mysql+pymysql://inventory_user:password@localhost/inventory_db?charset=utf8mb4
```

## 起動方法

```powershell
.\.venv\Scripts\python app.py
```

起動後、ブラウザで [http://127.0.0.1:5000](http://127.0.0.1:5000) を開いてください。

コマンドを毎回打ちたくない場合は、[start_tool.bat](</C:/Users/Owner/Desktop/script/自動引当/start_tool.bat>) をダブルクリックすると、サーバー起動とブラウザ表示までまとめて実行できます。

## 便利コマンド

デモデータ投入:

```powershell
$env:FLASK_APP='app.py'
.\.venv\Scripts\flask seed-demo
```

引当再計算:

```powershell
$env:FLASK_APP='app.py'
.\.venv\Scripts\flask recalculate
```

漏れチェック:

```powershell
$env:FLASK_APP='app.py'
.\.venv\Scripts\flask run-checks
```

## 画面一覧

1. ダッシュボード
2. 受注・発送状況
3. EMS梱包リスト
4. 4段階漏れチェック
5. 引当実行
6. 日本在庫管理
7. アラート一覧

## 実装メモ

- 受注・在庫の Yahoo API 取込は、認証情報未提供のため今回は接続点のみ確保し、運用前提の主要ロジックと UI を先に構築しています。
- SQLite でも動作しますが、`create_tables.sql` は MySQL 本番投入用に残しています。
- デモデータを入れると、即納引当、発注漏れ、遅延、EMS到着後の日本在庫仕分けまで確認できます。
