#!/usr/bin/env python3
import logging
import sys
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    logger = logging.getLogger("run_all_imports")

    try:
        from app import create_app
        from routes.import_data import run_all_imports_job

        logger.info("run_all_imports started")
        app = create_app()
        run_all_imports_job(app)
        logger.info("run_all_imports finished")
        return 0
    except Exception:
        logger.exception("run_all_imports failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
