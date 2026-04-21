"""Yahoo ショッピング ストアAPI クライアント
- リフレッシュトークンでアクセストークンを自動取得
- トークンローテーション（新リフレッシュトークン）を自動で .env に保存
- 以降は手動操作なしで永続的に使用可能
"""
import os, re, json, base64
import urllib.request, urllib.parse, urllib.error
from datetime import datetime, timedelta


YAHOO_TOKEN_URL        = 'https://auth.login.yahoo.co.jp/yconnect/v2/token'
YAHOO_ORDER_SEARCH_URL = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/searchOrder'
YAHOO_ORDER_INFO_URL   = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo'
YAHOO_STOCK_GET_URL    = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/getStock'
YAHOO_STOCK_SET_URL    = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/setStock'


class YahooAPI:

    def __init__(self):
        self.client_id     = os.getenv('YAHOO_CLIENT_ID', '')
        self.client_secret = os.getenv('YAHOO_CLIENT_SECRET', '')
        self.seller_id     = os.getenv('YAHOO_SELLER_ID', 'niyantarose')
        self.refresh_token = os.getenv('YAHOO_REFRESH_TOKEN', '')
        self._access_token = None

    # ── アクセストークン取得（自動リフレッシュ） ─────────────────────
    def get_access_token(self) -> str:
        """リフレッシュトークン → アクセストークン変換。
        ローテーションで新リフレッシュトークンが返ったら自動保存。"""
        if not self.refresh_token:
            raise ValueError('YAHOO_REFRESH_TOKEN が設定されていません。'
                             'http://localhost:5001/oauth/start から再認証してください。')

        credentials = base64.b64encode(
            f'{self.client_id}:{self.client_secret}'.encode()
        ).decode()

        data = urllib.parse.urlencode({
            'grant_type':    'refresh_token',
            'refresh_token': self.refresh_token,
        }).encode('utf-8')

        req = urllib.request.Request(
            YAHOO_TOKEN_URL,
            data=data,
            headers={
                'Authorization': f'Basic {credentials}',
                'Content-Type':  'application/x-www-form-urlencoded',
            }
        )

        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                token_data = json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'トークン取得失敗 HTTP {e.code}: {body}')

        access_token      = token_data.get('access_token', '')
        new_refresh_token = token_data.get('refresh_token', '')

        # ローテーション: 新しいリフレッシュトークンを自動保存
        if new_refresh_token and new_refresh_token != self.refresh_token:
            self._save_refresh_token(new_refresh_token)
            self.refresh_token = new_refresh_token
            os.environ['YAHOO_REFRESH_TOKEN'] = new_refresh_token

        if not access_token:
            raise RuntimeError(f'アクセストークンが取得できませんでした: {token_data}')

        self._access_token = access_token
        return access_token

    # ── 受注検索 ─────────────────────────────────────────────────────
    def search_orders(self, days: int = 7, start: int = 1, hits: int = 100) -> dict:
        """直近 days 日の受注を取得して辞書で返す。"""
        access_token = self.get_access_token()

        end_dt   = datetime.now()
        start_dt = end_dt - timedelta(days=days)

        params = {
            'sellerId':      self.seller_id,
            'orderTimeFrom': start_dt.strftime('%Y%m%d%H%M%S'),
            'orderTimeTo':   end_dt.strftime('%Y%m%d%H%M%S'),
            'condition':     '3',   # 全ステータス（元の動作確認済みパラメータ）
            'hits':          str(hits),
            'start':         str(start),
            'output':        'json',
        }

        url = YAHOO_ORDER_SEARCH_URL + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(
            url,
            headers={'Authorization': f'Bearer {access_token}'}
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'受注検索 HTTP {e.code}: {body}')

    def fetch_all_orders(self, days: int = 365) -> list:
        """全ページを取得してフラットなリストで返す。"""
        all_orders = []
        start = 1
        hits  = 100

        while True:
            result = self.search_orders(days=days, start=start, hits=hits)
            orders = self._extract_orders(result)
            if not orders:
                break
            all_orders.extend(orders)

            total = int(result.get('ResultSet', {}).get('@totalResultsAvailable',
                        result.get('TotalCount', 0)))
            if start + hits - 1 >= total:
                break
            start += hits

        return all_orders

    @staticmethod
    def _extract_orders(result: dict) -> list:
        """レスポンスから受注リストを抽出（APIバージョン差異を吸収）。"""
        # パターン1: ResultSet > Result > OrderInfo
        rs = result.get('ResultSet', {})
        if rs:
            result_block = rs.get('Result', {})
            orders = result_block.get('OrderInfo', [])
            if isinstance(orders, dict):
                orders = [orders]
            if orders:
                return orders

        # パターン2: OrderList
        orders = result.get('OrderList', {}).get('Order', [])
        if isinstance(orders, dict):
            orders = [orders]
        return orders or []

    # ── 受注明細取得 ─────────────────────────────────────────────────
    def get_order_detail(self, order_id: str) -> dict:
        """1件の受注の詳細（商品明細含む）を取得。"""
        access_token = self.get_access_token()

        params = {
            'sellerId': self.seller_id,
            'orderId':  order_id,
            'output':   'json',
        }
        url = YAHOO_ORDER_INFO_URL + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(
            url,
            headers={'Authorization': f'Bearer {access_token}'}
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'受注詳細 HTTP {e.code}: {body}')

    # ── 商品・在庫 一括ダウンロード ──────────────────────────────────
    DOWNLOAD_REQUEST_URL = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/downloadRequest'
    DOWNLOAD_SUBMIT_URL  = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/downloadSubmit'

    def _download_request(self, access_token: str, file_type: str) -> bool:
        """downloadRequest で Yahoo側のファイル生成をリクエストする。
        type: '1'=商品, '2'=在庫"""
        import time
        from xml.etree import ElementTree as ET
        data = urllib.parse.urlencode({
            'seller_id': self.seller_id,
            'type':      file_type,
        }).encode('utf-8')
        req = urllib.request.Request(
            self.DOWNLOAD_REQUEST_URL, data=data,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type':  'application/x-www-form-urlencoded',
            }
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read().decode('utf-8')
            root   = ET.fromstring(body)
            status = (root.findtext('.//Status') or '').strip().upper()
            return status == 'OK'
        except Exception as e:
            import logging
            logging.warning(f'downloadRequest type={file_type} エラー: {e}')
            return False

    @staticmethod
    def _decode_csv_bytes(raw: bytes) -> str:
        """CSVバイト列をエンコーディング自動検出でデコードする（UTF-8-SIG, UTF-8, CP932 を試みる）。"""
        for enc in ('utf-8-sig', 'utf-8', 'cp932', 'shift_jis'):
            try:
                return raw.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return raw.decode('utf-8', errors='replace')

    def _download_submit_with_retry(self, access_token: str, file_type: str,
                                     max_wait: int = 120) -> str:
        """downloadSubmit をポーリングで取得（pm-08003 が返る間は最大max_wait秒リトライ）。
        返り値: CSVテキスト（デコード済み）"""
        import time
        interval  = 10   # リトライ間隔（秒）
        elapsed   = 0
        params = {'seller_id': self.seller_id, 'type': file_type}
        url    = self.DOWNLOAD_SUBMIT_URL + '?' + urllib.parse.urlencode(params)

        while elapsed <= max_wait:
            req = urllib.request.Request(url, headers={'Authorization': f'Bearer {access_token}'})
            try:
                with urllib.request.urlopen(req, timeout=120) as resp:
                    raw = resp.read()
                return self._decode_csv_bytes(raw)
            except urllib.error.HTTPError as e:
                body = e.read().decode('utf-8', errors='replace')
                if 'pm-08003' in body and elapsed < max_wait:
                    import logging
                    logging.info(f'downloadSubmit type={file_type}: ファイル未生成（{elapsed}秒経過）、{interval}秒後リトライ')
                    time.sleep(interval)
                    elapsed += interval
                else:
                    raise RuntimeError(f'downloadSubmit(type={file_type}) HTTP {e.code}: {body}')

        raise RuntimeError(f'downloadSubmit: {max_wait}秒待ってもファイルが生成されませんでした')

    def download_inventory_csv(self) -> list:
        """Yahoo在庫データを一括ダウンロード（type=2）。
        downloadRequest → ポーリングdownloadSubmit の2ステップで取得。
        返り値: [{'code': 'item-01', 'sub_code': 'color1', 'quantity': 1000}, ...]
        quantityがNoneの場合は在庫無限大。"""
        import csv, io
        access_token = self.get_access_token()

        # Step 1: ファイル生成リクエスト
        self._download_request(access_token, '2')

        # Step 2: ファイルが準備できるまでポーリング（最大120秒）
        content = self._download_submit_with_retry(access_token, '2', max_wait=120)

        items = []
        reader = csv.DictReader(io.StringIO(content))
        for row in reader:
            code     = (row.get('code') or '').strip()
            if not code:
                continue
            qty_str  = (row.get('quantity') or '').strip()
            try:
                qty = int(qty_str) if qty_str else None
            except ValueError:
                qty = None
            try:
                allow_overdraft = int((row.get('allow-overdraft') or '0').strip())
            except ValueError:
                allow_overdraft = 0
            items.append({
                'code':            code,
                'sub_code':        (row.get('sub-code') or '').strip(),
                'quantity':        qty,             # None = 在庫無限大
                'allow_overdraft': allow_overdraft, # 1=注文取り/お取り寄せ, 0=実在庫
            })
        return items

    def download_item_csv(self) -> list:
        """Yahoo商品データを一括ダウンロード（type=1）。商品名・価格を取得。
        downloadRequest → ポーリングdownloadSubmit の2ステップで取得。
        返り値: [{'code': ..., 'name': ..., 'price': ...}, ...]"""
        import csv, io
        access_token = self.get_access_token()

        # Step 1: ファイル生成リクエスト
        self._download_request(access_token, '1')

        # Step 2: ポーリングでダウンロード
        content = self._download_submit_with_retry(access_token, '1', max_wait=120)

        items = []
        reader = csv.DictReader(io.StringIO(content))
        for row in reader:
            # Yahoo商品CSVの列名は環境によって異なるため複数候補を試みる
            def g(*keys):
                for k in keys:
                    v = (row.get(k) or '').strip()
                    if v:
                        return v
                return ''
            code = g('item-code', 'code', '商品コード')
            if not code:
                continue
            try:
                price = int(str(g('price', '販売価格', 'selling-price')).replace(',', '') or '0')
            except ValueError:
                price = 0
            items.append({
                'code':  code,
                'name':  g('item-name', 'name', '商品名'),
                'price': price,
            })
        return items

    # ── 在庫取得（単品・バッチ対応） ─────────────────────────────────
    def get_stock(self, item_code: str, sub_code: str = '') -> dict:
        """1件の在庫を取得（後方互換ラッパー）"""
        code = f'{item_code}:{sub_code}' if sub_code else item_code
        results = self.get_stock_batch([code])
        return results[0] if results else {}

    def get_stock_batch(self, item_codes: list) -> list:
        """最大1,000件の在庫を一括取得。レスポンスはXMLを解析して返す。
        item_codes: ['itemCode', 'itemCode:subCode', ...] 形式"""
        from xml.etree import ElementTree as ET
        access_token = self.get_access_token()

        codes_str = ','.join(str(c).strip() for c in item_codes[:1000])
        data = urllib.parse.urlencode({
            'seller_id': self.seller_id,
            'item_code': codes_str,
        }).encode('utf-8')
        req = urllib.request.Request(
            YAHOO_STOCK_GET_URL,
            data=data,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type':  'application/x-www-form-urlencoded',
            }
        )
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                content = resp.read().decode('utf-8')
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'getStock HTTP {e.code}: {body}')

        root = ET.fromstring(content)
        results = []
        for result in root.findall('Result'):
            qty_text = (result.findtext('Quantity') or '').strip()
            results.append({
                'item_code': (result.findtext('ItemCode') or '').strip(),
                'sub_code':  (result.findtext('SubCode')  or '').strip(),
                'quantity':  int(qty_text) if qty_text else None,  # None=無限大
                'status':    (result.findtext('Status')   or '0').strip(),
            })
        return results

    def get_all_stock_batch(self, item_codes: list) -> list:
        """全商品コードを1,000件ずつバッチでgetStock（5万商品対応）"""
        import time
        all_results = []
        batch_size  = 1000
        for i in range(0, len(item_codes), batch_size):
            batch = item_codes[i:i + batch_size]
            try:
                results = self.get_stock_batch(batch)
                all_results.extend(results)
            except Exception as e:
                import logging
                logging.warning(f'getStock batch {i}～{i+batch_size} エラー: {e}')
            if i + batch_size < len(item_codes):
                time.sleep(0.5)  # レート制限対策
        return all_results

    # ── 在庫更新（単品・バッチ対応） ─────────────────────────────────
    def set_stock(self, item_code: str, quantity: int, sub_code: str = '') -> dict:
        """1件の在庫を更新"""
        from xml.etree import ElementTree as ET
        access_token = self.get_access_token()
        code = f'{item_code}:{sub_code}' if sub_code else item_code
        data = urllib.parse.urlencode({
            'seller_id': self.seller_id,
            'item_code': code,
            'quantity':  str(quantity),
        }).encode('utf-8')
        req = urllib.request.Request(
            YAHOO_STOCK_SET_URL, data=data,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type':  'application/x-www-form-urlencoded',
            }
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                content = resp.read().decode('utf-8')
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'setStock HTTP {e.code}: {body}')
        root   = ET.fromstring(content)
        result = root.find('Result')
        if result is not None:
            return {
                'item_code': (result.findtext('ItemCode') or '').strip(),
                'sub_code':  (result.findtext('SubCode')  or '').strip(),
                'quantity':  (result.findtext('Quantity') or '').strip(),
            }
        return {}

    def upload_stock_file(self, csv_path: str) -> bool:
        """在庫CSVファイルをYahooにアップロード（在庫アップロードAPI）
        ファイル形式: https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadStockFile"""
        from xml.etree import ElementTree as ET
        try:
            import requests as req_lib
        except ImportError:
            raise ImportError('requests が必要です: pip install requests')
        access_token = self.get_access_token()
        url = (f'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadStockFile'
               f'?seller_id={urllib.parse.quote(self.seller_id)}')
        with open(csv_path, 'rb') as f:
            resp = req_lib.post(
                url,
                headers={'Authorization': f'Bearer {access_token}'},
                files={'file': (os.path.basename(csv_path), f, 'text/csv')},
                timeout=120,
            )
        root   = ET.fromstring(resp.content)
        status = (root.findtext('Status') or '').strip()
        return status.upper() == 'OK'

    # ── .env 保存 ────────────────────────────────────────────────────
    @staticmethod
    def _save_refresh_token(token: str):
        env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env')
        try:
            with open(env_path, 'r', encoding='utf-8') as f:
                content = f.read()
            content = re.sub(
                r'^YAHOO_REFRESH_TOKEN=.*$',
                f'YAHOO_REFRESH_TOKEN={token}',
                content,
                flags=re.MULTILINE
            )
            with open(env_path, 'w', encoding='utf-8') as f:
                f.write(content)
        except Exception as e:
            import logging
            logging.warning(f'リフレッシュトークン保存失敗: {e}')
