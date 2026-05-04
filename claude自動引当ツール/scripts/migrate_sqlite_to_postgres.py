#!/usr/bin/env python3
"""
SQLite（読み取り専用）から PostgreSQL へデータを複製するテスト用スクリプト。

本番 SQLite は書き換えない。--sqlite-path にコピー先を渡すことを推奨。
PostgreSQL 接続は --env-file（既定: リポジトリ直下の .env.pgtest）の USE_SQLITE / DATABASE_URL 等。
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
import traceback
from pathlib import Path
from typing import Any, Iterable

# ─── 投入順（FK 依存） ─────────────────────────────────────────────────────
MODEL_ORDER: list = []


def _sqlite_uri_readonly(abs_path: Path) -> str:
    u = abs_path.expanduser().resolve().as_uri()
    return f'{u}?mode=ro'


def _connect_sqlite_readonly(path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(_sqlite_uri_readonly(path), uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _run_scalar(conn: sqlite3.Connection, sql: str, params: tuple = ()) -> Any:
    cur = conn.execute(sql, params)
    row = cur.fetchone()
    return row[0] if row else None


def run_precheck_sqlite(sqlite_path: Path) -> int:
    """重複・孤立 FK を SQLite 上で検査。問題があれば 1。"""
    conn = _connect_sqlite_readonly(sqlite_path)
    try:
        issues: list[str] = []

        def chk(name: str, sql: str) -> int:
            n = int(_run_scalar(conn, sql) or 0)
            if n:
                issues.append(f'  [{name}] {n} 件')
            return n

        chk(
            'inventory 重複 (product_code, product_sub_code, inventory_type)',
            """
            SELECT COUNT(*) FROM (
              SELECT 1 FROM inventory
              GROUP BY product_code, COALESCE(product_sub_code, ''), inventory_type
              HAVING COUNT(*) > 1
            )
            """,
        )
        chk(
            'orders.yahoo_order_id 重複',
            """
            SELECT COUNT(*) FROM (
              SELECT 1 FROM orders GROUP BY yahoo_order_id HAVING COUNT(*) > 1
            )
            """,
        )
        chk(
            'ems.ems_number 重複',
            """
            SELECT COUNT(*) FROM (
              SELECT 1 FROM ems GROUP BY ems_number HAVING COUNT(*) > 1
            )
            """,
        )
        chk(
            'import_logs.filename 重複',
            """
            SELECT COUNT(*) FROM (
              SELECT 1 FROM import_logs GROUP BY filename HAVING COUNT(*) > 1
            )
            """,
        )
        chk(
            '孤立 order_items (order_id)',
            """
            SELECT COUNT(*) FROM order_items oi
            LEFT JOIN orders o ON o.id = oi.order_id
            WHERE o.id IS NULL
            """,
        )
        chk(
            '孤立 ems_items (ems_id)',
            """
            SELECT COUNT(*) FROM ems_items ei
            LEFT JOIN ems e ON e.id = ei.ems_id
            WHERE e.id IS NULL
            """,
        )
        chk(
            '孤立 ems_items (order_item_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM ems_items ei
            LEFT JOIN order_items oi ON oi.id = ei.order_item_id
            WHERE ei.order_item_id IS NOT NULL AND oi.id IS NULL
            """,
        )
        chk(
            '孤立 purchases (order_item_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM purchases p
            LEFT JOIN order_items oi ON oi.id = p.order_item_id
            WHERE p.order_item_id IS NOT NULL AND oi.id IS NULL
            """,
        )
        chk(
            '孤立 allocations (order_item_id)',
            """
            SELECT COUNT(*) FROM allocations a
            LEFT JOIN order_items oi ON oi.id = a.order_item_id
            WHERE oi.id IS NULL
            """,
        )
        chk(
            '孤立 allocations (ems_item_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM allocations a
            LEFT JOIN ems_items ei ON ei.id = a.ems_item_id
            WHERE a.ems_item_id IS NOT NULL AND ei.id IS NULL
            """,
        )
        chk(
            '孤立 alerts (order_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM alerts al
            LEFT JOIN orders o ON o.id = al.order_id
            WHERE al.order_id IS NOT NULL AND o.id IS NULL
            """,
        )
        chk(
            '孤立 alerts (order_item_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM alerts al
            LEFT JOIN order_items oi ON oi.id = al.order_item_id
            WHERE al.order_item_id IS NOT NULL AND oi.id IS NULL
            """,
        )
        chk(
            '孤立 japan_inventory_staging (ems_item_id)',
            """
            SELECT COUNT(*) FROM japan_inventory_staging j
            LEFT JOIN ems_items ei ON ei.id = j.ems_item_id
            WHERE ei.id IS NULL
            """,
        )
        chk(
            '孤立 japan_inventory_staging (assigned_order_item_id が非NULLだが存在しない)',
            """
            SELECT COUNT(*) FROM japan_inventory_staging j
            LEFT JOIN order_items oi ON oi.id = j.assigned_order_item_id
            WHERE j.assigned_order_item_id IS NOT NULL AND oi.id IS NULL
            """,
        )

        if issues:
            print('事前チェック: 問題あり (exit 1)')
            print('\n'.join(issues))
            return 1
        print('事前チェック: 問題なし')
        return 0
    finally:
        conn.close()


def _table_counts_sqlite(conn: sqlite3.Connection, tables: Iterable[str]) -> dict[str, int]:
    out: dict[str, int] = {}
    for t in tables:
        out[t] = int(_run_scalar(conn, f'SELECT COUNT(*) FROM {t}'))
    return out


def _row_to_mapping(row: sqlite3.Row, model) -> dict[str, Any]:
    cols = {c.key for c in model.__table__.columns}
    return {k: row[k] for k in row.keys() if k in cols}


def _load_models():
    global MODEL_ORDER
    from models.order import Order
    from models.order_item import OrderItem
    from models.ems import Ems
    from models.ems_item import EmsItem
    from models.inventory import Inventory
    from models.import_log import ImportLog
    from models.purchase import Purchase
    from models.allocation import Allocation
    from models.alert import Alert
    from models.japan_inventory import JapanInventoryStaging

    MODEL_ORDER = [
        Order,
        OrderItem,
        Ems,
        EmsItem,
        Inventory,
        ImportLog,
        Purchase,
        Allocation,
        Alert,
        JapanInventoryStaging,
    ]


TRUNCATE_TABLES_SQL = """
TRUNCATE TABLE
  japan_inventory_staging,
  alerts,
  allocations,
  purchases,
  import_logs,
  inventory,
  ems_items,
  order_items,
  ems,
  orders
RESTART IDENTITY CASCADE
"""


def migrate(sqlite_path: Path, env_file: Path, truncate: bool) -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    from dotenv import load_dotenv

    if not env_file.is_file():
        print(f'環境ファイルが見つかりません: {env_file}', file=sys.stderr)
        return 1

    load_dotenv(env_file, override=True)
    os.environ['USE_SQLITE'] = 'false'

    # app モジュール import 時に create_app() が走るため、上記 env 設定の直後に import する
    import app as app_module

    _load_models()
    from models import db
    from sqlalchemy import text

    table_names = [m.__tablename__ for m in MODEL_ORDER]

    sqlite_conn = _connect_sqlite_readonly(sqlite_path)
    try:
        src_counts = _table_counts_sqlite(sqlite_conn, table_names)
        print('SQLite 件数:', src_counts)
    finally:
        sqlite_conn.close()

    flask_app = app_module.app

    try:
        with flask_app.app_context():
            db.create_all()

            nonempty = []
            for m in MODEL_ORDER:
                c = db.session.query(m).count()
                if c > 0:
                    nonempty.append(m.__tablename__)

            if nonempty and not truncate:
                print(
                    'PostgreSQL に既存データがあります:',
                    nonempty,
                    '  --truncate を付けるか手動で空にしてください。',
                    file=sys.stderr,
                )
                return 1

            with db.session.begin():
                if truncate and nonempty:
                    db.session.execute(text(TRUNCATE_TABLES_SQL))

                sl_conn = _connect_sqlite_readonly(sqlite_path)
                try:
                    for model in MODEL_ORDER:
                        tbl = model.__tablename__
                        rows = sl_conn.execute(f'SELECT * FROM {tbl}').fetchall()
                        maps = [_row_to_mapping(r, model) for r in rows]
                        if maps:
                            db.session.bulk_insert_mappings(model, maps)
                        print(f'  投入 {tbl}: {len(maps)} 件')

                    for tbl in table_names:
                        db.session.execute(
                            text(
                                f"""
                                SELECT setval(
                                    pg_get_serial_sequence('{tbl}', 'id'),
                                    COALESCE((SELECT MAX(id) FROM {tbl}), 1),
                                    (SELECT MAX(id) FROM {tbl}) IS NOT NULL
                                )
                                """
                            )
                        )
                        print(f'  setval: {tbl}')
                finally:
                    sl_conn.close()

            dst_counts = {}
            for m in MODEL_ORDER:
                tbl = m.__tablename__
                dst_counts[tbl] = db.session.query(m).count()

        print('PostgreSQL 件数:', dst_counts)
        for t in table_names:
            s, d = src_counts[t], dst_counts[t]
            if s != d:
                print(f'件数不一致: {t} sqlite={s} pg={d}', file=sys.stderr)
                return 1
        print('件数一致を確認しました。')
        return 0

    except Exception:
        print(traceback.format_exc(), file=sys.stderr)
        try:
            with flask_app.app_context():
                db.session.rollback()
        except Exception:
            pass
        return 1


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    default_env = repo_root / '.env.pgtest'

    p = argparse.ArgumentParser(description='SQLite → PostgreSQL テスト移行')
    p.add_argument(
        '--sqlite-path',
        default=os.getenv('SQLITE_SOURCE_PATH', ''),
        help='読み取り元 SQLite ファイル（環境変数 SQLITE_SOURCE_PATH でも可）',
    )
    p.add_argument(
        '--env-file',
        default=str(default_env),
        help=f'PostgreSQL 接続用 .env（既定: {default_env}）',
    )
    p.add_argument(
        '--precheck-only',
        action='store_true',
        help='SQLite の事前チェックのみ実行して終了',
    )
    p.add_argument(
        '--truncate',
        action='store_true',
        help='PostgreSQL 側を TRUNCATE ... CASCADE してから投入（テスト DB のみで使用）',
    )
    args = p.parse_args()

    if not args.sqlite_path:
        print('--sqlite-path または SQLITE_SOURCE_PATH を指定してください。', file=sys.stderr)
        return 1

    sqlite_path = Path(args.sqlite_path)
    if not sqlite_path.is_file():
        print(f'SQLite ファイルがありません: {sqlite_path}', file=sys.stderr)
        return 1

    env_file = Path(args.env_file)

    if args.precheck_only:
        return run_precheck_sqlite(sqlite_path)

    return migrate(sqlite_path, env_file, args.truncate)


if __name__ == '__main__':
    raise SystemExit(main())
