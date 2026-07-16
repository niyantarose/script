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

# orders の String(n) 相当カラムと PostgreSQL 側の想定最大長（models/order.py と一致させる）
ORDERS_STRING_MAX_LEN: dict[str, int] = {
    'yahoo_order_id': 100,
    'customer_name': 255,
    'yahoo_ship_status': 50,
    'yahoo_order_status': 10,
    'yahoo_pay_status': 50,
    'status': 50,
    'ship_name': 255,
}

# 事前チェックで最大文字数のみ報告（Text 型・制限なし）
ORDERS_TEXT_LENGTH_INFO: tuple[str, ...] = (
    'ship_company_code',
    'gift_wrap_message',
    'delay_memo',
)


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


def _sqlite_orders_column_names(conn: sqlite3.Connection) -> set[str]:
    cur = conn.execute("SELECT name FROM pragma_table_info('orders')")
    return {str(r[0]) for r in cur.fetchall()}


def _check_orders_string_lengths(conn: sqlite3.Connection) -> tuple[list[str], list[str]]:
    """(エラー行, 情報行) を返す。エラー行があれば precheck は失敗させる。"""
    err_lines: list[str] = []
    info_lines: list[str] = []
    try:
        cols = _sqlite_orders_column_names(conn)
    except sqlite3.Error:
        return err_lines, info_lines
    if not cols:
        return err_lines, info_lines

    ref_col = 'yahoo_order_id' if 'yahoo_order_id' in cols else 'id'

    for col, max_len in ORDERS_STRING_MAX_LEN.items():
        if col not in cols:
            continue
        mx = _run_scalar(
            conn, f"SELECT MAX(LENGTH(COALESCE({col}, ''))) FROM orders"
        )
        mx_i = int(mx or 0)
        info_lines.append(f'  [orders.{col}] SQLite 最大文字数={mx_i}（PG 想定上限={max_len}）')
        if mx_i > max_len:
            row = conn.execute(
                f"SELECT {ref_col}, SUBSTR(COALESCE({col}, ''), 1, 120) FROM orders "
                f"WHERE LENGTH(COALESCE({col}, '')) = ? LIMIT 1",
                (mx_i,),
            ).fetchone()
            sample = f' ref={row[0]!r} 先頭120字={row[1]!r}' if row else ''
            err_lines.append(
                f'  [orders.{col}] 最大 {mx_i} 文字 > PG 想定 {max_len}{sample}'
            )

    for col in ORDERS_TEXT_LENGTH_INFO:
        if col not in cols:
            continue
        mx = _run_scalar(
            conn, f"SELECT MAX(LENGTH(COALESCE({col}, ''))) FROM orders"
        )
        mx_i = int(mx or 0)
        info_lines.append(f'  [orders.{col}] SQLite 最大文字数={mx_i}（PG は TEXT）')

    return err_lines, info_lines


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
            '孤立 allocations (order_item_id)',
            """
            SELECT COUNT(*) FROM allocations a
            LEFT JOIN order_items oi ON oi.id = a.order_item_id
            WHERE oi.id IS NULL
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
        len_err, len_info = _check_orders_string_lengths(conn)
        issues.extend(len_err)

        if issues:
            print('事前チェック: 問題あり (exit 1)')
            print('\n'.join(issues))
            return 1
        print('事前チェック: 問題なし')
        if len_info:
            print('--- orders 文字数（参考） ---')
            print('\n'.join(len_info))
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


def _order_row_from_sqlite(row: sqlite3.Row, model) -> dict[str, Any]:
    """orders のみ: SQLite に列が無い場合はデフォルト、is_seen は 0/1 を bool に。"""
    d = _row_to_mapping(row, model)
    if 'yahoo_pay_status' not in d or d.get('yahoo_pay_status') is None:
        d['yahoo_pay_status'] = ''
    else:
        d['yahoo_pay_status'] = str(d['yahoo_pay_status'])

    if 'is_seen' not in d or d.get('is_seen') is None:
        d['is_seen'] = False
    else:
        v = d['is_seen']
        if isinstance(v, bool):
            d['is_seen'] = v
        elif isinstance(v, (int, float)):
            d['is_seen'] = bool(int(v))
        elif isinstance(v, (bytes, bytearray)):
            d['is_seen'] = bool(v[0]) if len(v) else False
        elif isinstance(v, str):
            s = v.strip().lower()
            d['is_seen'] = s in ('1', 'true', 'yes', 't')
        else:
            d['is_seen'] = bool(v)
    return d


def _sqlite_row_to_mapping(row: sqlite3.Row, model) -> dict[str, Any]:
    if model.__tablename__ == 'orders':
        return _order_row_from_sqlite(row, model)
    return _row_to_mapping(row, model)


def _load_models():
    global MODEL_ORDER
    from models.order import Order
    from models.order_item import OrderItem
    from models.inventory import Inventory
    from models.import_log import ImportLog
    from models.allocation import Allocation
    from models.alert import Alert
    from models.stock_transaction import StockTransaction
    from models.mall_sku import MallSku

    MODEL_ORDER = [
        Order,
        OrderItem,
        Inventory,
        ImportLog,
        Allocation,
        Alert,
        StockTransaction,
        MallSku,
    ]


TRUNCATE_TABLES_SQL = """
TRUNCATE TABLE
  mall_skus,
  stock_transactions,
  alerts,
  allocations,
  import_logs,
  inventory,
  order_items,
  orders
RESTART IDENTITY CASCADE
"""


def _reset_db_session(db) -> None:
    """count / create_all 後に残る暗黙トランザクションを捨て、新規トランザクションを開始できる状態にする。"""
    try:
        db.session.rollback()
    except Exception:
        pass
    try:
        db.session.remove()
    except Exception:
        pass


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
            _reset_db_session(db)

            # ── 事前カウント（このブロックのトランザクションは投入前に切り離す） ──
            nonempty = []
            for m in MODEL_ORDER:
                c = db.session.query(m).count()
                if c > 0:
                    nonempty.append(m.__tablename__)

            if nonempty and not truncate:
                _reset_db_session(db)
                print(
                    'PostgreSQL に既存データがあります:',
                    nonempty,
                    '  --truncate を付けるか手動で空にしてください。',
                    file=sys.stderr,
                )
                return 1

            _reset_db_session(db)

            # ── TRUNCATE + INSERT + setval を単一トランザクションで実行 ──
            with db.session.begin():
                if truncate and nonempty:
                    db.session.execute(text(TRUNCATE_TABLES_SQL))

                sl_conn = _connect_sqlite_readonly(sqlite_path)
                try:
                    for model in MODEL_ORDER:
                        tbl = model.__tablename__
                        rows = sl_conn.execute(f'SELECT * FROM {tbl}').fetchall()
                        maps = [_sqlite_row_to_mapping(r, model) for r in rows]
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

            _reset_db_session(db)

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
                _reset_db_session(db)
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
