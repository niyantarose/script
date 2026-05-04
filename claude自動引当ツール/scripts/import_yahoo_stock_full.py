#!/usr/bin/env python3
"""Yahoo在庫フル同期を Gunicorn 経由せず CLI で実行する。"""
import logging
import sys
import traceback
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    log = logging.getLogger("import_yahoo_stock_full")

    try:
        from app import create_app
        from routes.import_data import run_yahoo_stock_full_import

        app = create_app()
        log.info("Yahoo在庫フル同期 CLI 開始")
        with app.app_context():
            try:
                result = run_yahoo_stock_full_import()
            except Exception:
                from models import db

                db.session.rollback()
                raise

        if result.get("status") != "ok":
            log.error("異常終了: status=%s result=%s", result.get("status"), result)
            return 1

        log.info(
            "Yahoo在庫フル同期 CLI 終了 status=%s imported=%s updated=%s total=%s message=%s",
            result.get("status"),
            result.get("imported"),
            result.get("updated"),
            result.get("total"),
            result.get("message"),
        )
        return 0
    except Exception:
        log.error("Yahoo在庫フル同期 CLI 失敗:\n%s", traceback.format_exc())
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
