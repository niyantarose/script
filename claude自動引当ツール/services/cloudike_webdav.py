"""Cloudike WebDAV 連携 - 発注/EMS ファイルの自動ダウンロード"""
import os
import re
import requests
from requests.auth import HTTPBasicAuth
from xml.etree import ElementTree


class CloudikeWebDAV:
    # .env で上書き可能。デフォルトは webdav.cloudike.com
    BASE_URL       = os.getenv('CLOUDIKE_BASE_URL',       'https://webdav.cloudike.com')
    PURCHASE_FOLDER = os.getenv('CLOUDIKE_PURCHASE_FOLDER', '/05.와타나베/01.와타나베주문/')
    EMS_FOLDER      = os.getenv('CLOUDIKE_EMS_FOLDER',      '/05.와타나베/02.와타나베발송리스트/')
    PURCHASE_PATTERN = r'Watanabe_list_(\d{6}).*\.xlsm$'
    # EMSファイル名は韓国語（발송리스트）
    EMS_PATTERN = r'EMS\ubc1c\uc1a1\ub9ac\uc2a4\ud2b8_(\d{6}).*\.xlsx$'

    def __init__(self):
        self.username = os.getenv('CLOUDIKE_USERNAME', '')
        self.password = os.getenv('CLOUDIKE_PASSWORD', '')
        self.download_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'downloads')
        os.makedirs(self.download_dir, exist_ok=True)

    def _auth(self):
        return HTTPBasicAuth(self.username, self.password)

    def list_files(self, folder):
        """フォルダ内のファイル一覧を PROPFIND で取得"""
        url = f'{self.BASE_URL}{folder}'
        resp = requests.request(
            'PROPFIND', url, auth=self._auth(),
            headers={'Depth': '1'}, timeout=30
        )
        resp.raise_for_status()

        files = []
        root = ElementTree.fromstring(resp.content)
        ns = {'d': 'DAV:'}
        for response in root.findall('d:response', ns):
            href = response.find('d:href', ns)
            if href is not None and href.text and not href.text.endswith('/'):
                filename = href.text.split('/')[-1]
                files.append({'name': filename, 'href': href.text})
        return files

    def find_latest_file(self, folder, pattern):
        """パターンにマッチするファイルのうち YYMMDD が最新のものを返す"""
        files = self.list_files(folder)
        matched = []
        for f in files:
            m = re.search(pattern, f['name'])
            if m:
                date_str = m.group(1)  # YYMMDD
                matched.append((date_str, f))
        if not matched:
            return None
        matched.sort(key=lambda x: x[0], reverse=True)
        return matched[0][1]  # {'name': ..., 'href': ...}

    def download_file(self, remote_href, local_filename=None):
        """ファイルをダウンロードしてローカルパスを返す"""
        url = f'{self.BASE_URL}{remote_href}'
        resp = requests.get(url, auth=self._auth(), timeout=60)
        resp.raise_for_status()

        if not local_filename:
            local_filename = remote_href.split('/')[-1]
        local_path = os.path.join(self.download_dir, local_filename)
        with open(local_path, 'wb') as f:
            f.write(resp.content)
        return local_path

    def list_all_purchase_files(self):
        """発注フォルダの全ファイル一覧（パターン一致）を日付昇順で返す"""
        return self._list_all(self.PURCHASE_FOLDER, self.PURCHASE_PATTERN)

    def list_all_ems_files(self):
        """EMSフォルダの全ファイル一覧（パターン一致）を日付昇順で返す"""
        return self._list_all(self.EMS_FOLDER, self.EMS_PATTERN)

    def _list_all(self, folder, pattern):
        """フォルダ内のパターン一致ファイルを日付昇順で返す"""
        try:
            files = self.list_files(folder)
        except Exception:
            return []
        matched = []
        for f in files:
            m = re.search(pattern, f['name'])
            if m:
                matched.append((m.group(1), f))  # (YYMMDD, file_info)
        matched.sort(key=lambda x: x[0])
        return [f for _, f in matched]

    def download_latest_purchase(self):
        """最新の発注ファイルをダウンロード（後方互換用）"""
        files = self.list_all_purchase_files()
        if not files:
            return None, None
        f = files[-1]  # 最新
        return self.download_file(f['href'], f['name']), f['name']

    def download_latest_ems(self):
        """最新のEMSファイルをダウンロード（後方互換用）"""
        files = self.list_all_ems_files()
        if not files:
            return None, None
        f = files[-1]  # 最新
        return self.download_file(f['href'], f['name']), f['name']
