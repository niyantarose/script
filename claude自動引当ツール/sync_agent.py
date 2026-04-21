"""
Google Drive → VPS 自動同期エージェント
=========================================
Google Drive フォルダを定期監視して、新しい/更新されたExcelファイルを
VPS の在庫引当ツールに自動アップロードします。

使い方:
  python sync_agent.py              # 1回だけ実行
  python sync_agent.py --loop       # 15分ごとに繰り返し実行
  python sync_agent.py --server     # HTTPトリガーサーバー起動（ブラウザの更新ボタン用）
                                    # ポート5050でリッスン、定期同期も同時実行

タスクスケジューラへの登録（推奨: --server モード）:
  タスク名: zaiko-sync-agent
  プログラム: pythonw
  引数: sync_agent.py --server
  開始場所: スクリプトのフォルダを指定
"""

import os, re, time, logging, hashlib, json, threading
from datetime import datetime
from pathlib import Path

# ── 設定（.env から読み込み） ──────────────────────────────────────────
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / '.env')

VPS_URL          = os.getenv('VPS_URL', 'http://133.167.89.53:5000')
WATCH_PURCHASE   = os.getenv('WATCH_FOLDER_PURCHASE', '')
WATCH_EMS        = os.getenv('WATCH_FOLDER_EMS', '')
LOCAL_SYNC_PORT  = int(os.getenv('LOCAL_SYNC_PORT', 5050))
PURCHASE_PATTERN = re.compile(r'Watanabe_list_\d{6}.*\.xlsm?$', re.IGNORECASE)
EMS_PATTERN      = re.compile(r'EMS(?:発送リスト|발송리스트)_\d{6}.*\.xlsx?$', re.IGNORECASE)

# 状態ファイル（アップロード済みファイルのハッシュを保存）
STATE_FILE = Path(__file__).parent / '.sync_state'

# ── ログ設定 ──────────────────────────────────────────────────────────
# ログファイル: 書き込み可能なパスを順番に試みる
_log_candidates = [
    Path(r'C:\zaiko\sync_agent.log'),            # Task Scheduler / pythonw 実行
    Path(__file__).parent / 'sync_agent.log',    # ターミナル直接起動
]
_log_file = None
for _p in _log_candidates:
    try:
        _p.parent.mkdir(parents=True, exist_ok=True)
        _p.open('a', encoding='utf-8').close()
        _log_file = _p
        break
    except (PermissionError, OSError):
        continue

# コンソールが使える場合のみ StreamHandler を追加（pythonw では不要）
_has_console = False
try:
    import sys as _sys
    if _sys.stdout is not None:
        _sys.stdout.fileno()  # コンソールなし（pythonw）なら例外が出る
        _has_console = True
except Exception:
    pass

_handlers = []
if _has_console:
    _handlers.append(logging.StreamHandler())
if _log_file:
    _handlers.append(logging.FileHandler(_log_file, encoding='utf-8'))
if not _handlers:
    _handlers.append(logging.NullHandler())

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=_handlers,
)
log = logging.getLogger(__name__)
log.info(f'ログファイル: {_log_file or "(なし)"}')


# ── ユーティリティ ────────────────────────────────────────────────────

def file_hash(path):
    """ファイルの MD5 ハッシュ（変更検出用）"""
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def load_state():
    """アップロード済みファイルの状態を読み込む {filename: hash}"""
    state = {}
    if STATE_FILE.exists():
        for line in STATE_FILE.read_text(encoding='utf-8').splitlines():
            parts = line.strip().split('\t')
            if len(parts) == 2:
                state[parts[0]] = parts[1]
    return state


def save_state(state):
    """状態ファイルに書き込む"""
    lines = [f'{k}\t{v}' for k, v in state.items()]
    STATE_FILE.write_text('\n'.join(lines), encoding='utf-8')


def upload_file(local_path, file_type):
    """VPS の取込 API にファイルをアップロード
    file_type: 'purchase' または 'ems'
    """
    try:
        import requests
        url = f'{VPS_URL}/import/upload/daniel_{file_type}'
        filename = Path(local_path).name
        log.info(f'  📤 アップロード中: {url}')
        with open(local_path, 'rb') as f:
            resp = requests.post(
                url,
                files={'file': (filename, f, 'application/octet-stream')},
                timeout=120
            )
        log.info(f'  HTTP {resp.status_code}  (body={len(resp.content)}bytes)')
        if not resp.content:
            log.error(f'  ❌ {filename}: VPSから空のレスポンス (HTTP {resp.status_code})')
            log.error(f'     VPSサービスが起動しているか確認してください: sudo systemctl status zaiko-tool')
            return False
        try:
            data = resp.json()
        except Exception:
            log.error(f'  ❌ {filename}: JSONパース失敗 (HTTP {resp.status_code})')
            log.error(f'     レスポンス: {resp.text[:300]}')
            return False
        if data.get('status') == 'ok':
            log.info(f'  ✅ {filename} → {data.get("imported", 0)}件取込')
            return True
        else:
            log.error(f'  ❌ {filename} エラー: {data.get("message")}')
            return False
    except Exception as e:
        log.error(f'  ❌ アップロード失敗 [{Path(local_path).name}]: {e}')
        return False


def scan_and_upload(folder, pattern, file_type, state):
    """フォルダをスキャンして変更があればアップロード"""
    if not folder or not os.path.isdir(folder):
        return

    for fname in sorted(os.listdir(folder)):
        if not pattern.search(fname):
            continue
        fpath = os.path.join(folder, fname)
        if not os.path.isfile(fpath):
            continue

        current_hash = file_hash(fpath)
        if state.get(fname) == current_hash:
            continue  # 変更なし

        log.info(f'📄 新規/更新ファイル検出: {fname}')
        if upload_file(fpath, file_type):
            state[fname] = current_hash


# ── メイン ────────────────────────────────────────────────────────────

def run_once():
    """1回のスキャンを実行"""
    log.info('=== 同期開始 ===')
    state = load_state()

    # 発注リスト
    if WATCH_PURCHASE:
        log.info(f'発注フォルダ: {WATCH_PURCHASE}')
        scan_and_upload(WATCH_PURCHASE, PURCHASE_PATTERN, 'purchase', state)
    else:
        log.warning('WATCH_FOLDER_PURCHASE が未設定')

    # EMSリスト
    if WATCH_EMS:
        log.info(f'EMS フォルダ: {WATCH_EMS}')
        scan_and_upload(WATCH_EMS, EMS_PATTERN, 'ems', state)
    else:
        log.warning('WATCH_FOLDER_EMS が未設定')

    save_state(state)
    log.info('=== 同期完了 ===')


def run_loop(interval_minutes=15):
    """定期実行ループ（タスクスケジューラを使わない場合）"""
    log.info(f'🚀 同期エージェント起動 (間隔: {interval_minutes}分)')
    while True:
        try:
            run_once()
        except Exception as e:
            log.error(f'エラー: {e}')
        log.info(f'次回まで {interval_minutes} 分待機...')
        time.sleep(interval_minutes * 60)


def start_trigger_server(interval_minutes=15):
    """ブラウザの「今すぐ取込」ボタンから呼べるローカルHTTPサーバー。
    CORS対応。port 5050（LOCAL_SYNC_PORT）でリッスン。
    バックグラウンドスレッドで定期同期も実行。
    """
    from http.server import HTTPServer, BaseHTTPRequestHandler

    class SyncHandler(BaseHTTPRequestHandler):
        def _cors(self):
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type')

        def do_OPTIONS(self):
            self.send_response(200)
            self._cors()
            self.end_headers()

        def do_GET(self):
            # ヘルスチェック用
            if self.path == '/health':
                body = json.dumps({'ok': True, 'agent': 'zaiko-sync'}).encode()
                self.send_response(200)
                self._cors()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(body)
            else:
                self.send_response(404)
                self.end_headers()

        def do_POST(self):
            if self.path == '/sync':
                log.info('🔘 ブラウザからの手動同期トリガー受信')
                try:
                    state = load_state()
                    imported_p = 0
                    imported_e = 0

                    if WATCH_PURCHASE and os.path.isdir(WATCH_PURCHASE):
                        before = len(state)
                        scan_and_upload(WATCH_PURCHASE, PURCHASE_PATTERN, 'purchase', state)
                        imported_p = len(state) - before

                    if WATCH_EMS and os.path.isdir(WATCH_EMS):
                        before = len(state)
                        scan_and_upload(WATCH_EMS, EMS_PATTERN, 'ems', state)
                        imported_e = len(state) - before

                    save_state(state)

                    parts = []
                    if imported_p: parts.append(f'発注{imported_p}件')
                    if imported_e: parts.append(f'EMS{imported_e}件')
                    msg = '新規取込: ' + ' / '.join(parts) if parts else '新しいファイルなし'
                    body = json.dumps({'ok': True, 'message': msg,
                                       'purchase': imported_p, 'ems': imported_e},
                                      ensure_ascii=False).encode('utf-8')
                    self.send_response(200)
                    self._cors()
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(body)
                except Exception as ex:
                    log.error(f'手動同期エラー: {ex}')
                    body = json.dumps({'ok': False, 'message': str(ex)},
                                      ensure_ascii=False).encode('utf-8')
                    self.send_response(500)
                    self._cors()
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(body)
            else:
                self.send_response(404)
                self.end_headers()

        def log_message(self, *args):
            pass  # HTTPサーバーのコンソールログを抑制

    # 定期同期をバックグラウンドスレッドで実行
    def _periodic():
        log.info(f'🔄 定期同期スレッド開始 ({interval_minutes}分間隔)')
        while True:
            try:
                run_once()
            except Exception as ex:
                log.error(f'定期同期エラー: {ex}')
            time.sleep(interval_minutes * 60)

    t = threading.Thread(target=_periodic, daemon=True)
    t.start()

    server = HTTPServer(('127.0.0.1', LOCAL_SYNC_PORT), SyncHandler)
    log.info(f'✅ 同期トリガーサーバー起動: http://127.0.0.1:{LOCAL_SYNC_PORT}/')
    log.info(f'   ブラウザの「今すぐ取込」ボタンから呼び出せます')
    server.serve_forever()


if __name__ == '__main__':
    import sys
    args = sys.argv[1:]
    if '--server' in args:
        # サーバーモード: HTTPトリガー + 定期同期（推奨）
        start_trigger_server(interval_minutes=15)
    elif '--loop' in args:
        # 連続実行モード（HTTPなし）
        run_loop(interval_minutes=15)
    else:
        # 1回実行モード（タスクスケジューラ単発実行用）
        run_once()
