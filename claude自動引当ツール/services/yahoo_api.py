"""Yahoo ストアクリエイター API 連携"""
import os
import requests
from datetime import datetime


class YahooAPI:
    TOKEN_URL = 'https://auth.login.yahoo.co.jp/yconnect/v2/token'
    ORDER_API = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderList'
    ORDER_INFO_API = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo'
    STOCK_API = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/getStock'
    SET_STOCK_API = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/setStock'

    def __init__(self):
        self.client_id = os.getenv('YAHOO_CLIENT_ID', '')
        self.client_secret = os.getenv('YAHOO_CLIENT_SECRET', '')
        self.refresh_token = os.getenv('YAHOO_REFRESH_TOKEN', '')
        self.access_token = None

    def _refresh_access_token(self):
        """リフレッシュトークンでアクセストークンを更新"""
        resp = requests.post(self.TOKEN_URL, data={
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.refresh_token,
        })
        resp.raise_for_status()
        data = resp.json()
        self.access_token = data['access_token']
        return self.access_token

    def _get_headers(self):
        if not self.access_token:
            self._refresh_access_token()
        return {'Authorization': f'Bearer {self.access_token}'}

    def fetch_orders(self, seller_id, **kwargs):
        """受注リストを取得"""
        params = {
            'sellerId': seller_id,
            'Condition': kwargs.get('condition', 1),  # 1=新規受注
            'DateType': kwargs.get('date_type', 1),
            'StartDate': kwargs.get('start_date', ''),
            'EndDate': kwargs.get('end_date', ''),
        }
        resp = requests.get(self.ORDER_API, headers=self._get_headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    def fetch_order_info(self, seller_id, order_id):
        """受注詳細を取得"""
        params = {
            'sellerId': seller_id,
            'Target': 'OrderId',
            'Value': order_id,
        }
        resp = requests.get(self.ORDER_INFO_API, headers=self._get_headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    def fetch_stock(self, seller_id, item_code):
        """在庫数を取得"""
        params = {
            'sellerId': seller_id,
            'itemCode': item_code,
        }
        resp = requests.get(self.STOCK_API, headers=self._get_headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    def set_stock(self, seller_id, item_code, quantity, sub_code=None):
        """在庫数を更新（加算）"""
        data = {
            'sellerId': seller_id,
            'itemCode': item_code,
            'quantity': quantity,
        }
        if sub_code:
            data['subCode'] = sub_code
        resp = requests.post(self.SET_STOCK_API, headers=self._get_headers(), data=data)
        resp.raise_for_status()
        return resp.json()
