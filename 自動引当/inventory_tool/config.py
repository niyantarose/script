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
    YAHOO_STORE_ACCOUNT = os.getenv("YAHOO_STORE_ACCOUNT", "")
    YAHOO_STORE_ID = os.getenv("YAHOO_STORE_ID", "")
    GOOGLE_SHEETS_FILE_ID = os.getenv("GOOGLE_SHEETS_FILE_ID", "")
    GOOGLE_DRIVE_CREDENTIALS_JSON = os.getenv("GOOGLE_DRIVE_CREDENTIALS_JSON", "")
    GOOGLE_SHEETS_PURCHASE_SHEET = os.getenv("GOOGLE_SHEETS_PURCHASE_SHEET", "発注リスト")
    GOOGLE_SHEETS_EMS_SHEET = os.getenv("GOOGLE_SHEETS_EMS_SHEET", "EMSリスト")
    CLOUDIKE_WEBDAV_URL = os.getenv("CLOUDIKE_WEBDAV_URL", "https://webdav.cloudike.com")
    CLOUDIKE_WEBDAV_USERNAME = os.getenv("CLOUDIKE_WEBDAV_USERNAME", "")
    CLOUDIKE_WEBDAV_PASSWORD = os.getenv("CLOUDIKE_WEBDAV_PASSWORD", "")
    CLOUDIKE_DANIEL_PURCHASE_DIR = os.getenv("CLOUDIKE_DANIEL_PURCHASE_DIR", "/05.와타나베/01.와타나베주문/")
    CLOUDIKE_DANIEL_EMS_DIR = os.getenv("CLOUDIKE_DANIEL_EMS_DIR", "/05.와타나베/02.와타나베발송리스트/")
