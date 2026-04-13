from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "inventory-tool-dev-secret")
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL",
        f"sqlite:///{(BASE_DIR / 'inventory_tool.db').as_posix()}",
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    DELAY_WARNING_DAYS = int(os.getenv("DELAY_WARNING_DAYS", "5"))
    EMS_LEAD_DAYS = int(os.getenv("EMS_LEAD_DAYS", "2"))
    SHIPPING_BUFFER_DAYS = int(os.getenv("SHIPPING_BUFFER_DAYS", "2"))
    YAHOO_CLIENT_ID = os.getenv("YAHOO_CLIENT_ID", "")
    YAHOO_CLIENT_SECRET = os.getenv("YAHOO_CLIENT_SECRET", "")
    YAHOO_REFRESH_TOKEN = os.getenv("YAHOO_REFRESH_TOKEN", "")
