"""Google Sheets API 連携 - 発注リストとEMSリストの取得"""
import os
import json
import requests
from datetime import datetime


class GoogleSheetsAPI:
    SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets'

    def __init__(self):
        self.api_key = os.getenv('GOOGLE_API_KEY', '')
        self.spreadsheet_id = os.getenv('GOOGLE_SPREADSHEET_ID', '')
        # サービスアカウント認証を使う場合
        self.credentials_path = os.getenv('GOOGLE_CREDENTIALS_PATH', '')
        self.access_token = None

    def _get_access_token(self):
        """サービスアカウントでアクセストークンを取得"""
        if self.credentials_path and os.path.exists(self.credentials_path):
            # google-auth ライブラリを使用
            try:
                from google.oauth2 import service_account
                from google.auth.transport.requests import Request

                creds = service_account.Credentials.from_service_account_file(
                    self.credentials_path,
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
                creds.refresh(Request())
                self.access_token = creds.token
                return self.access_token
            except ImportError:
                pass
        return None

    def _get_headers(self):
        if self.access_token:
            return {'Authorization': f'Bearer {self.access_token}'}
        return {}

    def _build_url(self, range_str):
        url = f'{self.SHEETS_API}/{self.spreadsheet_id}/values/{range_str}'
        if self.api_key and not self.access_token:
            url += f'?key={self.api_key}'
        return url

    def fetch_sheet(self, sheet_name, range_str='A:Z'):
        """指定シートのデータを取得"""
        self._get_access_token()
        full_range = f'{sheet_name}!{range_str}'
        url = self._build_url(full_range)
        resp = requests.get(url, headers=self._get_headers())
        resp.raise_for_status()
        data = resp.json()
        return data.get('values', [])

    def fetch_purchases(self):
        """発注リストシートを取得してパース"""
        rows = self.fetch_sheet('発注リスト')
        if len(rows) < 2:
            return []

        headers = rows[0]
        purchases = []
        for row in rows[1:]:
            item = {}
            for i, h in enumerate(headers):
                item[h] = row[i] if i < len(row) else ''
            purchases.append(item)
        return purchases

    def fetch_ems_list(self):
        """EMSリストシートを取得してパース"""
        rows = self.fetch_sheet('EMSリスト')
        if len(rows) < 2:
            return []

        headers = rows[0]
        ems_list = []
        for row in rows[1:]:
            item = {}
            for i, h in enumerate(headers):
                item[h] = row[i] if i < len(row) else ''
            ems_list.append(item)
        return ems_list
