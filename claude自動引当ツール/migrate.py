"""
DB マイグレーションスクリプト
実行: python migrate.py

変更内容:
  purchases  : order_item_id を nullable に / purchase_no・order_id カラム追加
  ems        : purchase_no・order_id カラム追加
  ems_items  : order_item_id を nullable に / product_name カラム追加
"""
import sqlite3
import os
import sys

# SQLite ファイルのパスを特定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
candidates = [
    os.path.join(BASE_DIR, 'instance', 'inventory.db'),
    os.path.join(BASE_DIR, 'inventory.db'),
]
db_path = None
for c in candidates:
    if os.path.exists(c):
        db_path = c
        break

if not db_path:
    print('ERROR: inventory.db が見つかりません')
    print('確認パス:', candidates)
    sys.exit(1)

print(f'DB: {db_path}')
conn = sqlite3.connect(db_path)
conn.execute('PRAGMA foreign_keys = OFF')
cur = conn.cursor()


def col_names(table):
    cur.execute(f'PRAGMA table_info({table})')
    return {row[1] for row in cur.fetchall()}


# ── purchases テーブル ────────────────────────────────────────────────
cols = col_names('purchases')
if 'purchase_no' not in cols or 'order_id' not in cols:
    print('purchases: カラム追加 & order_item_id を nullable に変更...')
    cur.executescript("""
        ALTER TABLE purchases RENAME TO purchases_old;

        CREATE TABLE purchases (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_no      VARCHAR(100),
            order_id         VARCHAR(100),
            order_item_id    INTEGER REFERENCES order_items(id),
            product_code     VARCHAR(100) NOT NULL,
            product_sub_code VARCHAR(100),
            product_name     VARCHAR(200),
            quantity         INTEGER NOT NULL DEFAULT 1,
            shop_name        VARCHAR(100),
            ordered_at       DATE,
            status           VARCHAR(30) DEFAULT 'ordered',
            agent            VARCHAR(20) DEFAULT 'daniel',
            memo             TEXT,
            created_at       DATETIME,
            updated_at       DATETIME
        );

        INSERT INTO purchases (
            id, order_item_id, product_code, product_sub_code,
            product_name, quantity, shop_name, ordered_at,
            status, agent, memo, created_at, updated_at
        )
        SELECT id, order_item_id, product_code, product_sub_code,
               product_name, quantity, shop_name, ordered_at,
               status, agent, memo, created_at, updated_at
        FROM purchases_old;

        DROP TABLE purchases_old;
    """)
    print('  purchases: 完了')
else:
    print('purchases: スキップ（既に最新）')

# ── ems テーブル ──────────────────────────────────────────────────────
cols = col_names('ems')
if 'purchase_no' not in cols:
    cur.execute('ALTER TABLE ems ADD COLUMN purchase_no VARCHAR(100)')
    print('ems: purchase_no 追加')
if 'order_id' not in cols:
    cur.execute('ALTER TABLE ems ADD COLUMN order_id VARCHAR(100)')
    print('ems: order_id 追加')

# ── ems_items テーブル ────────────────────────────────────────────────
cols = col_names('ems_items')
if 'product_name' not in cols or 'purchase_no' not in cols or 'purchase_date' not in cols:
    print('ems_items: カラム追加（purchase_date / purchase_no / product_name）...')
    cur.executescript("""
        ALTER TABLE ems_items RENAME TO ems_items_old;

        CREATE TABLE ems_items (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            ems_id           INTEGER NOT NULL REFERENCES ems(id),
            order_item_id    INTEGER REFERENCES order_items(id),
            purchase_date    VARCHAR(50),
            purchase_no      VARCHAR(100),
            product_code     VARCHAR(100) NOT NULL,
            product_sub_code VARCHAR(100),
            product_name     VARCHAR(200),
            quantity         INTEGER NOT NULL DEFAULT 1,
            created_at       DATETIME
        );

        INSERT INTO ems_items (
            id, ems_id, order_item_id, purchase_no,
            product_code, product_sub_code, product_name, quantity, created_at
        )
        SELECT id, ems_id, order_item_id,
               CASE WHEN typeof(purchase_no)='text' THEN purchase_no ELSE NULL END,
               product_code, product_sub_code, product_name, quantity, created_at
        FROM ems_items_old;

        DROP TABLE ems_items_old;
    """)
    print('  ems_items: 完了')
else:
    print('ems_items: スキップ（既に最新）')

conn.commit()
conn.close()
print('\nマイグレーション完了！')
print('次のステップ: sudo systemctl restart zaiko-tool')
