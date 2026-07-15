#!/usr/bin/env python3
"""台湾CNシート × Yahoo!ショッピング実店舗 の商品コード突き合わせ（読み取り専用）。

シート内部の健全性チェックでは「シートの中の矛盾」しか見えないため、
実際に店に出ている商品(約10万件)とシートを商品コードで照合し、
作品IDずれ事故の実害や取りこぼしをレポートする。

- Yahoo側: 商品CSV一括ダウンロードAPI（呼び出し数回のみ、件数に比例しない）
- シート側: 閲覧リンクのCSVエクスポート（無認証・読み取りのみ）
- 出力: reports/recon_YYYYMMDD_HHMMSS.csv とコンソール要約。書き込みは一切しない

実行: python scripts/yahoo_sheet_recon.py [--max-wait 600] [--skip-yahoo FILE.csv]
"""
import argparse
import csv
import difflib
import io
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

# 台湾CNスプレッドシート（商品登録よろしく_台湾CNなど）
SPREADSHEET_ID = '1OSoZnNTMHrH5YgU-j7zwOsyQsjS6bcwjfZxc3kzeKWM'

# シート定義: (シート名, gid, コード列候補, タイトル列候補)
# コード列はYahooテンプレ送信と同じ優先順（まんが=親コード、他=商品コード（SKU））。
# 親コードが空の行（旧フロー時代の行など）はSKU（自動）に行単位でフォールバックする。
SHEET_DEFS = [
    ('台湾まんが',     '334744058',  ['親コード', 'SKU（自動）', 'SKU(自動)'], ['タイトル']),
    ('台湾書籍その他', '1871073917', ['商品コード（SKU）', '商品コード(SKU)', 'SKU（自動）', 'SKU(自動)'], ['タイトル']),
    ('台湾グッズ',     '1788781147', ['商品コード（SKU）', '商品コード(SKU)'], ['商品名（出品用）']),
    ('台湾雑誌',       '595657005',  ['商品コード（SKU）', '商品コード(SKU)'], ['商品名（出品用）']),
]

# 照合②（Yahooのみ）の対象スコープ: この形状のコードだけを「このシートの管轄」とみなす。
# 店舗全体には韓国商品・旧システム商品など管轄外が大量にあるため。
TAIWAN_CODE_RE = re.compile(r'^(TW|CN|HK)[A-Z]*\d{4}-')

# タイトル類似度がこの値未満なら「③タイトル不一致」として報告
TITLE_RATIO_THRESHOLD = 0.8

_DASHES_RE = re.compile(r'[‐‑‒–—―−ー]')
_SPACES_RE = re.compile(r'[\s　]+')


def normalize_code(value: str) -> str:
    """商品コードの表記揺れを吸収（Yahoo側は小文字化される・全角/ダッシュ揺れ対策）。"""
    s = unicodedata.normalize('NFKC', str(value or ''))
    s = _DASHES_RE.sub('-', s)
    s = _SPACES_RE.sub('', s)
    return s.upper().strip()


def normalize_title(value: str) -> str:
    s = unicodedata.normalize('NFKC', str(value or ''))
    return _SPACES_RE.sub(' ', s).strip()


def title_ratio(a: str, b: str) -> float:
    na, nb = normalize_title(a), normalize_title(b)
    if not na or not nb:
        return 1.0  # 比較不能は不一致扱いにしない
    return difflib.SequenceMatcher(None, na, nb).ratio()


def first_nonempty(row: dict, keys: list) -> str:
    for k in keys:
        v = (row.get(k) or '').strip()
        if v:
            return v
    return ''


def fetch_sheet_rows() -> list:
    """4シートをCSVエクスポートで取得し、行リストに正規化する。"""
    import requests
    rows = []
    for sheet_name, gid, code_cols, title_cols in SHEET_DEFS:
        url = (f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'
               f'/export?format=csv&gid={gid}')
        resp = requests.get(url, timeout=60)
        resp.raise_for_status()
        reader = csv.DictReader(io.StringIO(resp.content.decode('utf-8')))
        n = 0
        for i, row in enumerate(reader):
            code = first_nonempty(row, code_cols)
            if not code or code.upper().startswith('ERROR'):
                continue
            rows.append({
                'sheet': sheet_name,
                'row': i + 2,  # ヘッダー1行分
                'code': code,
                'title': first_nonempty(row, title_cols),
                'status': (row.get('登録状況') or '').strip(),
            })
            n += 1
        print(f'  {sheet_name}: コードあり {n} 行')
    return rows


def fetch_yahoo_items(max_wait: int) -> list:
    """Yahoo商品CSVを一括ダウンロード（読み取り専用）。"""
    from services.yahoo_api import YahooAPI
    api = YahooAPI()
    return api.download_item_csv(max_wait=max_wait)


def load_yahoo_items_from_csv(path: str) -> list:
    """--skip-yahoo 用: 保存済みの商品CSV（code,name列）から読み込む。"""
    items = []
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            code = (row.get('code') or row.get('item-code') or row.get('商品コード') or '').strip()
            if code:
                items.append({
                    'code': code,
                    'name': (row.get('name') or row.get('item-name') or row.get('商品名') or '').strip(),
                })
    return items


def reconcile(sheet_rows: list, yahoo_items: list) -> dict:
    """突き合わせ本体（純粋関数・テスト対象）。

    返り値: {
      'sheet_only':     ①シート登録済みなのにYahooに無い,
      'yahoo_only':     ②Yahooにあるのにシートに無い（台湾CN形状のみ・旧システム含む）,
      'title_mismatch': ③コード一致だがタイトル類似度が低い,
      'dup_in_sheet':   ④シート内で同一コードが複数行,
    }
    """
    yahoo_by_code = {}
    for item in yahoo_items:
        key = normalize_code(item.get('code'))
        if key:
            yahoo_by_code.setdefault(key, item)

    sheet_by_code = {}
    for r in sheet_rows:
        key = normalize_code(r['code'])
        if key:
            sheet_by_code.setdefault(key, []).append(r)

    buckets = {'sheet_only': [], 'yahoo_only': [], 'title_mismatch': [], 'dup_in_sheet': []}

    # ④ シート内コード重複（Yahooには同一コードを2つ置けないため、片方は未反映のはず）
    for key, rs in sheet_by_code.items():
        if len(rs) > 1:
            buckets['dup_in_sheet'].append({
                'code': key,
                'rows': [(r['sheet'], r['row'], r['title']) for r in rs],
            })

    # ①/③ シート起点
    for key, rs in sheet_by_code.items():
        registered = [r for r in rs if '登録済' in r['status']]
        yahoo_item = yahoo_by_code.get(key)
        if yahoo_item is None:
            if registered:
                r = registered[0]
                buckets['sheet_only'].append({
                    'code': key, 'sheet': r['sheet'], 'row': r['row'], 'title': r['title'],
                })
            continue
        # ③ タイトル照合（コードが店に実在する行のみ）
        r = registered[0] if registered else rs[0]
        ratio = title_ratio(r['title'], yahoo_item.get('name', ''))
        if ratio < TITLE_RATIO_THRESHOLD:
            buckets['title_mismatch'].append({
                'code': key, 'sheet': r['sheet'], 'row': r['row'],
                'sheet_title': r['title'], 'yahoo_name': yahoo_item.get('name', ''),
                'ratio': round(ratio, 3),
            })

    # ② Yahoo起点（台湾CN形状のみ）
    for key, item in yahoo_by_code.items():
        if TAIWAN_CODE_RE.match(key) and key not in sheet_by_code:
            buckets['yahoo_only'].append({'code': key, 'yahoo_name': item.get('name', '')})

    buckets['sheet_only'].sort(key=lambda x: x['code'])
    buckets['yahoo_only'].sort(key=lambda x: x['code'])
    buckets['title_mismatch'].sort(key=lambda x: x['ratio'])
    buckets['dup_in_sheet'].sort(key=lambda x: x['code'])
    return buckets


def write_report(buckets: dict, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f'recon_{datetime.now():%Y%m%d_%H%M%S}.csv'
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['区分', '商品コード', 'シート', '行', 'シート側タイトル', 'Yahoo商品名', '類似度'])
        for x in buckets['sheet_only']:
            w.writerow(['①シート登録済み・Yahoo無し', x['code'], x['sheet'], x['row'], x['title'], '', ''])
        for x in buckets['title_mismatch']:
            w.writerow(['③タイトル不一致', x['code'], x['sheet'], x['row'],
                        x['sheet_title'], x['yahoo_name'], x['ratio']])
        for x in buckets['dup_in_sheet']:
            locs = ' / '.join(f"{s}:{row}({t})" for s, row, t in x['rows'])
            w.writerow(['④シート内コード重複', x['code'], locs, '', '', '', ''])
        for x in buckets['yahoo_only']:
            w.writerow(['②Yahooのみ(旧システム含む)', x['code'], '', '', '', x['yahoo_name'], ''])
    return path


def main() -> int:
    # Windowsのcp932コンソールでも中国語タイトル等で落ちないようにする
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

    ap = argparse.ArgumentParser(description='台湾CNシート×Yahoo実店舗の商品コード照合（読み取り専用）')
    ap.add_argument('--max-wait', type=int, default=600,
                    help='Yahoo CSV生成のポーリング上限秒（10万件規模向けに既定600秒）')
    ap.add_argument('--skip-yahoo', metavar='FILE',
                    help='YahooAPIを呼ばず、保存済み商品CSV(code,name)を使う')
    ap.add_argument('--out', default=str(REPO_ROOT / 'reports'), help='レポート出力先ディレクトリ')
    args = ap.parse_args()

    from dotenv import load_dotenv
    load_dotenv(REPO_ROOT / '.env')

    print('■ シートCSV取得中...')
    sheet_rows = fetch_sheet_rows()

    if args.skip_yahoo:
        print(f'■ Yahoo商品データ: {args.skip_yahoo} から読込')
        yahoo_items = load_yahoo_items_from_csv(args.skip_yahoo)
    else:
        print('■ Yahoo商品CSVダウンロード中（ファイル生成待ちで数分かかることがあります）...')
        yahoo_items = fetch_yahoo_items(args.max_wait)
    print(f'  Yahoo商品: {len(yahoo_items)} 件')

    buckets = reconcile(sheet_rows, yahoo_items)
    path = write_report(buckets, Path(args.out))

    print('\n===== 照合結果 =====')
    print(f"① シート登録済みなのにYahooに無い : {len(buckets['sheet_only'])} 件"
          '（コード確定済みでも未出品の行は正常にここへ出ます）')
    print(f"② Yahooのみ(台湾CN形状・旧システム含む): {len(buckets['yahoo_only'])} 件")
    print(f"③ コード一致・タイトル不一致       : {len(buckets['title_mismatch'])} 件")
    print(f"④ シート内コード重複               : {len(buckets['dup_in_sheet'])} 件")
    for label, key, limit in (('①', 'sheet_only', 10), ('③', 'title_mismatch', 10), ('④', 'dup_in_sheet', 10)):
        for x in buckets[key][:limit]:
            print(f"  {label} {x['code']} : {x.get('title') or x.get('sheet_title') or ''}"
                  + (f" ⇔ {x['yahoo_name']} (類似度{x['ratio']})" if key == 'title_mismatch' else ''))
    print(f'\nレポート: {path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
