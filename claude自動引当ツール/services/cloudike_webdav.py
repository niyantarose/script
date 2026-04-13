"""Cloudike WebDAV 連携 - EMSファイルの自動ダウンロード"""
import os
import requests
from requests.auth import HTTPBasicAuth
from xml.etree import ElementTree


class CloudikeWebDAV:
    BASE_URL = 'https://webdav.cloudike.com'

    def __init__(self):
        self.username = os.getenv('CLOUDIKE_USERNAME', '')
        self.password = os.getenv('CLOUDIKE_PASSWORD', '')
        self.ems_folder = os.getenv('CLOUDIKE_EMS_FOLDER', '/EMS/')
        self.download_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'downloads')
        os.makedirs(self.download_dir, exist_ok=True)

    def _auth(self):
        return HTTPBasicAuth(self.username, self.password)

    def list_files(self, folder=None):
        """フォルダ内のファイル一覧を取得"""
        folder = folder or self.ems_folder
        url = f'{self.BASE_URL}{folder}'
        resp = requests.request('PROPFIND', url, auth=self._auth(), headers={'Depth': '1'})
        resp.raise_for_status()

        files = []
        root = ElementTree.fromstring(resp.content)
        ns = {'d': 'DAV:'}
        for response in root.findall('d:response', ns):
            href = response.find('d:href', ns)
            if href is not None and href.text and not href.text.endswith('/'):
                filename = href.text.split('/')[-1]
                files.append({
                    'name': filename,
                    'href': href.text,
                })
        return files

    def download_file(self, remote_path, local_filename=None):
        """ファイルをダウンロード"""
        url = f'{self.BASE_URL}{remote_path}'
        resp = requests.get(url, auth=self._auth())
        resp.raise_for_status()

        if not local_filename:
            local_filename = remote_path.split('/')[-1]

        local_path = os.path.join(self.download_dir, local_filename)
        with open(local_path, 'wb') as f:
            f.write(resp.content)

        return local_path

    def download_latest_ems(self):
        """最新のEMSファイルをダウンロード"""
        files = self.list_files()
        # Excelファイルのみフィルタ
        excel_files = [f for f in files if f['name'].endswith(('.xlsx', '.xls'))]
        if not excel_files:
            return None

        # 名前でソートして最新を取得
        excel_files.sort(key=lambda x: x['name'], reverse=True)
        latest = excel_files[0]
        return self.download_file(latest['href'])
