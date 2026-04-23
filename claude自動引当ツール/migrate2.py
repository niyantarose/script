"""
DB マイグレーション 2
実行: python migrate2.py

変更内容:
  inventory: product_name / yahoo_stock / price / location /
             is_immediate / last_synced_at カラム追加
"""
import sqlite3, os, sys

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
cur  = conn.cursor()


def col_names(table):
    cur.execute(f'PRAGMA table_info({table})')
    return {row[1] for row in cur.fetchall()}


# ── inventory テーブル ────────────────────────────────────────────────
cols = col_names('inventory')
new_cols = [
    ('product_name',   'VARCHAR(200)'),
    ('yahoo_stock',    'INTEGER DEFAULT 0'),
    ('price',          'INTEGER DEFAULT 0'),
    ('location',       'VARCHAR(100)'),
    ('is_immediate',   'BOOLEAN DEFAULT 0'),
    ('last_synced_at', 'DATETIME'),
]
for col, definition in new_cols:
    if col not in cols:
        cur.execute(f'ALTER TABLE inventory ADD COLUMN {col} {definition}')
        print(f'inventory: {col} 追加')
    else:
        print(f'inventory: {col} スキップ（既存）')

conn.commit()
conn.close()
print('\nマイグレーション完了！')
print('次のステップ: sudo systemctl restart zaiko-tool')
