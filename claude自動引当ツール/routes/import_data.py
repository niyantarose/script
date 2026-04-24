"""データ取込ルート"""
import os, tempfile
from flask import Blueprint, request, jsonify, current_app
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.ems import Ems
from models.ems_item import EmsItem
from models.inventory import Inventory
from models.import_log import ImportLog
from datetime import datetime, timedelta

bp = Blueprint('import_data', __name__, url_prefix='/import')


# ─── ユーティリティ ──────────────────────────────────────────────────────────

def _is_immediate(sub_code) -> bool:
    """サブコードの末尾文字で即納 / お取り寄せを判定（GASスクリプトと同ロジック）
    - 末尾が 'b'（大小文字不問）→ お取り寄せ（False）
    - 末尾が 'a' / サブコードなし / その他 → 即納（True）
    """
    sc = (sub_code or '').strip().lower()
    if not sc:
        return True      # サブコードなし → 即納
    return sc[-1] != 'b'


# ─── Yahoo ───────────────────────────────────────────────────────────────────

@bp.route('/yahoo_orders', methods=['POST'])
def import_yahoo_orders():
    """Yahoo受注を取込んで Order / OrderItem に保存"""
    try:
        from services.yahoo_api import YahooAPI
        body = request.get_json(silent=True) or {}
        days = int(body.get('days', 365))
        api = YahooAPI()
        orders = api.fetch_all_orders(days=days)

        imported_orders = 0
        imported_items  = 0
        new_item_ids    = []  # 自動引当用に新規明細IDを収集

        for order_raw in orders:
            cnt_o, cnt_i, ids = _upsert_yahoo_order(order_raw)
            imported_orders += cnt_o
            imported_items  += cnt_i
            new_item_ids.extend(ids)

        db.session.commit()

        # ── 新規明細を即座に自動引当 ──────────────────────────────
        alloc_stats = {'fully_allocated': 0, 'partial': 0, 'tori': 0,
                       'shortage': 0, 'skipped': 0}
        if new_item_ids:
            from services.allocation import run_auto_allocation
            alloc_stats = run_auto_allocation(item_ids=new_item_ids)
            db.session.commit()

        return jsonify({
            'status':          'ok',
            'imported_orders': imported_orders,
            'imported_items':  imported_items,
            'alloc':           alloc_stats,
            'message': (f'{imported_orders}件の受注・{imported_items}件の明細を取込みました'
                        f' | 引当: 即納{alloc_stats["fully_allocated"]}件'
                        f' / 部分{alloc_stats["partial"]}件'
                        f' / お取り寄せ{alloc_stats["tori"]}件'
                        f' / 不足{alloc_stats["shortage"]}件'),
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/yahoo_stock_diff', methods=['POST'])
def import_yahoo_stock_diff():
    """【Phase1: リアルタイム差分同期】
    Yahoo在庫をダウンロード → DBとの差分だけUPDATE（高速・低負荷）
    5分ごとのcron実行用。

    戦略:
      - 既存DBをメモリに1回だけロード（1 SELECT）
      - CSVと照合し、quantity/is_immediateが変わった商品のみUPDATE
      - 同じ値ならスキップ（DB書き込みなし）
      - 変更内容は最新50件のみレスポンス返却
    """
    try:
        from services.yahoo_api import YahooAPI
        api = YahooAPI()

        # ① 在庫CSV取得（type=2 のみ。差分同期では商品名は取らない=高速化）
        stock_items = api.download_inventory_csv()
        if not stock_items:
            return jsonify({'status': 'ok', 'changes': 0, 'message': 'Yahoo在庫データなし'})

        # ② 既存DB在庫をメモリにロード（1 SELECT）
        existing_map = {}
        for inv in Inventory.query.filter_by(inventory_type='yahoo').all():
            key = (inv.product_code, inv.product_sub_code or '')
            existing_map[key] = inv

        now       = datetime.now()
        new_cnt   = 0
        upd_cnt   = 0
        skip_cnt  = 0
        changes   = []   # 変更ログ（最新50件のみ保持）

        for item in stock_items:
            code     = item['code']
            sub_code = item['sub_code'] or None
            qty      = item['quantity']
            if not code or qty is None:
                continue

            # サブコード末尾が 'b' → お取り寄せ、それ以外 → 即納
            is_immediate = _is_immediate(sub_code)
            key = (code, sub_code or '')
            existing = existing_map.get(key)

            if existing:
                # ── 差分チェック: 変わった項目のみ更新 ──
                if existing.yahoo_stock != qty or existing.is_immediate != is_immediate:
                    if len(changes) < 50:
                        changes.append({
                            'code':    code,
                            'old_qty': existing.yahoo_stock,
                            'new_qty': qty,
                            'old_imm': existing.is_immediate,
                            'new_imm': is_immediate,
                        })
                    existing.yahoo_stock    = qty
                    existing.quantity       = qty
                    existing.is_immediate   = is_immediate
                    existing.last_synced_at = now
                    # available_qty を reserved_qty と突き合わせて再計算
                    existing.available_qty  = max(0, qty - (existing.reserved_qty or 0))
                    upd_cnt += 1
                else:
                    skip_cnt += 1  # 変更なし
            else:
                # 新規商品（Yahoo側で新規登録された）
                db.session.add(Inventory(
                    product_code=code,
                    product_sub_code=sub_code,
                    inventory_type='yahoo',
                    quantity=qty,
                    yahoo_stock=qty,
                    available_qty=qty,
                    is_immediate=is_immediate,
                    last_synced_at=now,
                ))
                new_cnt += 1

        db.session.commit()

        # ── 在庫が増えた商品に対し pending/shortage 明細を再引当 ──
        alloc_stats = {'fully_allocated': 0, 'partial': 0, 'tori': 0,
                       'shortage': 0, 'skipped': 0}
        if upd_cnt > 0 or new_cnt > 0:
            from services.allocation import run_auto_allocation
            alloc_stats = run_auto_allocation()   # pending/shortage 全件対象
            db.session.commit()

        return jsonify({
            'status':   'ok',
            'new':      new_cnt,
            'updated':  upd_cnt,
            'skipped':  skip_cnt,
            'total':    len(stock_items),
            'changes':  changes,
            'alloc':    alloc_stats,
            'message':  (f'差分同期完了: 新規{new_cnt} / 更新{upd_cnt} / 変更なし{skip_cnt}'
                         f'（合計{len(stock_items)}件チェック）'
                         f' | 引当: 即納{alloc_stats["fully_allocated"]}件'
                         f' / 部分{alloc_stats["partial"]}件'
                         f' / お取り寄せ{alloc_stats["tori"]}件'
                         f' / 不足{alloc_stats["shortage"]}件'),
        })
    except Exception as e:
        db.session.rollback()
        import traceback
        current_app.logger.error(f'yahoo_stock_diff エラー: {e}\n{traceback.format_exc()}')
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/yahoo_stock', methods=['POST'])
def import_yahoo_stock():
    """【初回・フル同期】Yahoo在庫＋商品名を一括取得して Inventory を差分更新。
    - 在庫CSV (type=2): quantity, allow_overdraft
    - 商品CSV (type=1): 商品名, 価格
    - allow_overdraft=0 かつ quantity>0 → is_immediate=True（即納）
    - allow_overdraft=1 → is_immediate=False（お取り寄せ）
    """
    try:
        from services.yahoo_api import YahooAPI
        api = YahooAPI()

        # ① 在庫CSV取得（type=2）
        stock_items = api.download_inventory_csv()
        if not stock_items:
            return jsonify({'status': 'ok', 'message': '在庫データがありませんでした', 'updated': 0})

        # ② 商品CSV取得（type=1）→ code をキーにした辞書を作成
        try:
            item_list  = api.download_item_csv()
            item_names = {i['code']: i for i in item_list}
        except Exception as e:
            current_app.logger.warning(f'商品CSV取得失敗（在庫のみ更新）: {e}')
            item_names = {}

        imported = 0
        updated  = 0
        now      = datetime.now()

        for item in stock_items:
            code     = item['code']
            sub_code = item['sub_code'] or None
            qty      = item['quantity']
            if not code or qty is None:
                continue  # コードなし or 在庫無限大はスキップ

            # サブコード末尾が 'b' → お取り寄せ、それ以外 → 即納
            is_immediate = _is_immediate(sub_code)

            # 商品名・価格（商品CSV から補完）
            item_info    = item_names.get(code, {})
            product_name = item_info.get('name', '')
            price        = item_info.get('price', 0)

            existing = Inventory.query.filter_by(
                product_code=code,
                product_sub_code=sub_code,
                inventory_type='yahoo',
            ).first()

            if existing:
                existing.yahoo_stock    = qty
                existing.quantity       = qty
                existing.is_immediate   = is_immediate
                existing.available_qty  = max(0, qty - (existing.reserved_qty or 0))
                existing.last_synced_at = now
                if product_name:
                    existing.product_name = product_name
                if price:
                    existing.price = price
                updated += 1
            else:
                db.session.add(Inventory(
                    product_code=code,
                    product_sub_code=sub_code,
                    product_name=product_name,
                    inventory_type='yahoo',
                    quantity=qty,
                    yahoo_stock=qty,
                    available_qty=qty,
                    price=price,
                    is_immediate=is_immediate,
                    last_synced_at=now,
                ))
                imported += 1

            if (imported + updated) % 500 == 0:
                db.session.flush()

        db.session.commit()

        # ── フル同期後も pending/shortage 明細を再引当 ──
        from services.allocation import run_auto_allocation
        alloc_stats = run_auto_allocation()
        db.session.commit()

        total = len(stock_items)
        immediate_cnt = sum(1 for i in stock_items if _is_immediate(i.get('sub_code', '')))
        return jsonify({
            'status':    'ok',
            'imported':  imported,
            'updated':   updated,
            'total':     total,
            'immediate': immediate_cnt,
            'alloc':     alloc_stats,
            'message':   (f'Yahoo在庫同期完了: {imported}件新規 / {updated}件更新（計{total}件）'
                          f' | 即納:{immediate_cnt}件'
                          f' | 引当: 即納{alloc_stats["fully_allocated"]}件'
                          f' / 部分{alloc_stats["partial"]}件'
                          f' / お取り寄せ{alloc_stats["tori"]}件'
                          f' / 不足{alloc_stats["shortage"]}件'),
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─── Yahoo CSV アップロード（5万商品一括登録）────────────────────────────────

@bp.route('/upload/yahoo_csv', methods=['POST'])
def upload_yahoo_csv():
    """Yahoo管理画面から出力した商品CSV/TSVをアップロードして Inventory に一括登録"""
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'ファイルが選択されていません'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'status': 'error', 'message': 'ファイル名が不明です'}), 400
    tmp = _save_upload(f)
    try:
        rows = _parse_yahoo_item_csv(tmp)
        imported, updated = _upsert_yahoo_inventory(rows)
        return jsonify({
            'status':   'ok',
            'imported': imported,
            'updated':  updated,
            'total':    len(rows),
            'message':  f'Yahoo在庫登録完了: {imported}件新規 / {updated}件更新（計{len(rows)}件）',
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        try: os.unlink(tmp)
        except: pass


def _parse_yahoo_item_csv(file_path):
    """Yahoo CSVを自動解析（UTF-8/Shift-JIS・タブ/カンマ区切り・日英ヘッダー対応）"""
    import csv

    encoding = 'utf-8-sig'
    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift-jis']:
        try:
            with open(file_path, 'r', encoding=enc) as fh:
                fh.read(4096)
            encoding = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue

    with open(file_path, 'r', encoding=encoding) as fh:
        sample = fh.read(2048)
    delimiter = '\t' if sample.count('\t') > sample.count(',') else ','

    items = []
    with open(file_path, 'r', encoding=encoding) as fh:
        reader = csv.DictReader(fh, delimiter=delimiter)
        for row in reader:
            def g(*keys):
                for k in keys:
                    v = (row.get(k) or '').strip()
                    if v:
                        return v
                return ''

            def gi(*keys):
                try:
                    return int(str(g(*keys)).replace(',', '') or '0')
                except ValueError:
                    return 0

            code = g('item-code', 'itemCode', '商品コード', 'item_code', 'sku')
            if not code:
                continue
            items.append({
                'product_code':     code,
                'product_sub_code': g('item-sub-code', 'subCode', 'サブコード'),
                'product_name':     g('item-name', 'itemName', '商品名', 'name'),
                'yahoo_stock':      gi('stock-quantity', 'quantity', 'stockQuantity', '在庫数'),
                'price':            gi('price', '販売価格', 'selling-price', 'unit-price'),
            })
    return items


def _upsert_yahoo_inventory(items):
    """Yahoo商品データを Inventory テーブルに差分upsert"""
    imported = 0
    updated  = 0
    now      = datetime.now()

    for item in items:
        product_code = item['product_code']
        sub_code     = item.get('product_sub_code') or None

        existing = Inventory.query.filter_by(
            product_code=product_code,
            product_sub_code=sub_code,
            inventory_type='yahoo',
        ).first()

        if existing:
            if item.get('product_name'):
                existing.product_name = item['product_name']
            existing.yahoo_stock    = item['yahoo_stock']
            existing.quantity       = item['yahoo_stock']
            if item.get('price'):
                existing.price = item['price']
            existing.last_synced_at = now
            updated += 1
        else:
            db.session.add(Inventory(
                product_code=product_code,
                product_sub_code=sub_code,
                product_name=item.get('product_name', ''),
                inventory_type='yahoo',
                quantity=item['yahoo_stock'],
                yahoo_stock=item['yahoo_stock'],
                price=item.get('price', 0),
                last_synced_at=now,
            ))
            imported += 1

        if (imported + updated) % 500 == 0:
            db.session.flush()

    db.session.commit()
    return imported, updated


# ─── Google Sheets ───────────────────────────────────────────────────────────

@bp.route('/google_purchases', methods=['POST'])
def import_google_purchases():
    try:
        from services.google_sheets import GoogleSheetsAPI
        api = GoogleSheetsAPI()
        rows = api.fetch_purchases()
        imported = _upsert_purchases(rows, agent='tegu')
        return jsonify({'status': 'ok', 'imported': imported})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/google_ems', methods=['POST'])
def import_google_ems():
    try:
        from services.google_sheets import GoogleSheetsAPI
        api = GoogleSheetsAPI()
        rows = api.fetch_ems_list()
        imported = _upsert_ems_rows(rows, agent='tegu')
        return jsonify({'status': 'ok', 'imported': imported})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─── Cloudike（ダニエル）────────────────────────────────────────────────────

@bp.route('/cloudike_purchase', methods=['POST'])
def import_cloudike_purchase():
    """Cloudike から最新の発注ファイル (Watanabe_list_YYMMDD.xlsm) を取込 → ダニエル発注リスト"""
    try:
        from services.cloudike_webdav import CloudikeWebDAV
        from services.purchase_excel_parser import PurchaseExcelParser

        webdav = CloudikeWebDAV()
        local_path, filename = webdav.download_latest_purchase()
        if not local_path:
            return jsonify({'status': 'ok', 'message': '発注ファイルが見つかりませんでした', 'imported': 0})

        parser = PurchaseExcelParser()
        rows = parser.parse(local_path)
        imported = _upsert_purchases_from_excel(rows, agent='daniel')

        log = ImportLog.query.filter_by(filename=filename).first()
        if log:
            log.record_count = imported
            log.imported_at  = datetime.now()
        else:
            db.session.add(ImportLog(filename=filename, file_type='purchase_daniel', record_count=imported))
        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported, 'filename': filename})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/cloudike_ems', methods=['POST'])
def import_cloudike_ems():
    """Cloudike から最新の EMS 発送リスト (EMS発送リスト_YYMMDD.xlsx) を取込 → ダニエルEMSリスト"""
    try:
        from services.cloudike_webdav import CloudikeWebDAV
        from services.ems_excel_parser import EmsExcelParser

        webdav = CloudikeWebDAV()
        local_path, filename = webdav.download_latest_ems()
        if not local_path:
            return jsonify({'status': 'ok', 'message': 'EMSファイルが見つかりませんでした', 'imported': 0})

        parser = EmsExcelParser()
        items = parser.parse(local_path)
        imported = _upsert_ems_items(items, agent='daniel')

        log = ImportLog.query.filter_by(filename=filename).first()
        if log:
            log.record_count = imported
            log.imported_at  = datetime.now()
        else:
            db.session.add(ImportLog(filename=filename, file_type='ems_daniel', record_count=imported))
        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported, 'filename': filename})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─── Google Drive（rclone 経由）────────────────────────────────────────────

GDRIVE_PURCHASE_FOLDER = 'gdrive:在庫引当/05.와타나베/01.와타나베주문'
GDRIVE_EMS_FOLDER      = 'gdrive:在庫引当/05.와타나베/02.와타나베발송리스트'


def _rclone_download_latest(remote_folder, pattern_re, suffix):
    """rclone で remote_folder の最新ファイルを tmp にダウンロードして返す。"""
    import subprocess, tempfile, re, os
    result = subprocess.run(
        ['/usr/bin/rclone', 'ls', remote_folder],
        capture_output=True, text=True, timeout=30
    )
    files = []
    for line in result.stdout.strip().splitlines():
        m = re.search(pattern_re, line)
        if m:
            files.append(m.group(0))
    if not files:
        return None, None
    latest = sorted(files)[-1]
    remote_path = remote_folder + '/' + latest
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    subprocess.run(['/usr/bin/rclone', 'copyto', remote_path, tmp_path],
                   capture_output=True, timeout=90, check=True)
    return tmp_path, latest


@bp.route('/gdrive_purchase', methods=['POST'])
def import_gdrive_purchase():
    """Google Drive から rclone で Watanabe_list_*.xlsm をダウンロードして取込"""
    import os
    try:
        tmp_path, filename = _rclone_download_latest(
            GDRIVE_PURCHASE_FOLDER,
            r'Watanabe_list_\d{6}.*\.xlsm?$',
            '.xlsm'
        )
        if not tmp_path:
            return jsonify({'status': 'ok', 'message': '発注ファイルが見つかりませんでした', 'imported': 0})

        from services.purchase_excel_parser import PurchaseExcelParser
        rows = PurchaseExcelParser().parse(tmp_path)
        imported = _upsert_purchases_from_excel(rows, agent='daniel')

        log = ImportLog.query.filter_by(filename=filename).first()
        if log:
            log.record_count = imported; log.imported_at = datetime.now()
        else:
            db.session.add(ImportLog(filename=filename, file_type='purchase_daniel', record_count=imported))
        db.session.commit()
        os.unlink(tmp_path)
        return jsonify({'status': 'ok', 'imported': imported, 'filename': filename,
                        'message': f'{filename} から {imported}件取込みました'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/gdrive_ems', methods=['POST'])
def import_gdrive_ems():
    """Google Drive から rclone で EMS発送リスト_*.xlsx をダウンロードして取込"""
    import os
    try:
        tmp_path, filename = _rclone_download_latest(
            GDRIVE_EMS_FOLDER,
            r'EMS(?:発送リスト|발송리스트)_\d{6}.*\.xlsx?$',
            '.xlsx'
        )
        if not tmp_path:
            return jsonify({'status': 'ok', 'message': 'EMSファイルが見つかりませんでした', 'imported': 0})

        from services.ems_excel_parser import EmsExcelParser
        items = EmsExcelParser().parse(tmp_path)
        imported = _upsert_ems_items(items, agent='daniel')

        log = ImportLog.query.filter_by(filename=filename).first()
        if log:
            log.record_count = imported; log.imported_at = datetime.now()
        else:
            db.session.add(ImportLog(filename=filename, file_type='ems_daniel', record_count=imported))
        db.session.commit()
        os.unlink(tmp_path)
        return jsonify({'status': 'ok', 'imported': imported, 'filename': filename,
                        'message': f'{filename} から {imported}件取込みました'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─── フォルダ監視取込（Google Drive / OneDrive 等） ──────────────────────────

PURCHASE_PATTERN_RE = r'Watanabe_list_\d{6}.*\.xlsm?$'
# 日本語名（EMS発送リスト_）と韓国語名（EMS발송리스트_）の両方に対応
EMS_PATTERN_RE = r'EMS(?:発送リスト|발송리스트)_\d{6}.*\.xlsx?$'


def _scan_folder(folder, pattern_re):
    """指定フォルダからパターンに合うファイル一覧を返す"""
    import re
    if not folder or not os.path.isdir(folder):
        return []
    result = []
    for fname in sorted(os.listdir(folder)):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath) and re.search(pattern_re, fname):
            result.append((fname, fpath))
    return result


def _folder_file_list(folder, pattern_re, file_type_label):
    files = []
    for fname, _ in _scan_folder(folder, pattern_re):
        log = ImportLog.query.filter_by(filename=fname).first()
        files.append({'name': fname, 'type': file_type_label,
                      'imported': log is not None,
                      'imported_at': log.imported_at.strftime('%m/%d %H:%M') if log else ''})
    return files


@bp.route('/folder_status', methods=['GET'])
def folder_status():
    """監視フォルダのファイル状況を返す"""
    kind = request.args.get('type', 'purchase')  # 'purchase' or 'ems'
    if kind == 'ems':
        folder = os.getenv('WATCH_FOLDER_EMS', '')
        files = _folder_file_list(folder, EMS_PATTERN_RE, 'EMS')
    else:
        folder = os.getenv('WATCH_FOLDER_PURCHASE', '')
        files = _folder_file_list(folder, PURCHASE_PATTERN_RE, '発注')
    return jsonify({'folder': folder, 'files': files,
                    'folder_exists': os.path.isdir(folder) if folder else False})


@bp.route('/from_folder', methods=['POST'])
def import_from_folder():
    """Google Drive フォルダから新しいファイルを取込"""
    kind = request.args.get('type', 'all')  # 'purchase', 'ems', 'all'
    results = {'purchase': [], 'ems': []}
    total = 0

    if kind in ('purchase', 'all'):
        from services.purchase_excel_parser import PurchaseExcelParser
        folder = os.getenv('WATCH_FOLDER_PURCHASE', '')
        for fname, fpath in _scan_folder(folder, PURCHASE_PATTERN_RE):
            if ImportLog.query.filter_by(filename=fname).first():
                results['purchase'].append({'file': fname, 'status': 'skip', 'count': 0})
                continue
            try:
                rows = PurchaseExcelParser().parse(fpath)
                cnt = _upsert_purchases_from_excel(rows, agent='daniel')
                db.session.add(ImportLog(filename=fname, file_type='purchase_daniel', record_count=cnt))
                db.session.commit()
                results['purchase'].append({'file': fname, 'status': 'ok', 'count': cnt})
                total += cnt
            except Exception as e:
                db.session.rollback()
                results['purchase'].append({'file': fname, 'status': 'error', 'message': str(e)})

    if kind in ('ems', 'all'):
        from services.ems_excel_parser import EmsExcelParser
        folder = os.getenv('WATCH_FOLDER_EMS', '')
        for fname, fpath in _scan_folder(folder, EMS_PATTERN_RE):
            if ImportLog.query.filter_by(filename=fname).first():
                results['ems'].append({'file': fname, 'status': 'skip', 'count': 0})
                continue
            try:
                items = EmsExcelParser().parse(fpath)
                cnt = _upsert_ems_items(items, agent='daniel')
                db.session.add(ImportLog(filename=fname, file_type='ems_daniel', record_count=cnt))
                db.session.commit()
                results['ems'].append({'file': fname, 'status': 'ok', 'count': cnt})
                total += cnt
            except Exception as e:
                db.session.rollback()
                results['ems'].append({'file': fname, 'status': 'error', 'message': str(e)})

    new_ok = [r for lst in results.values() for r in lst if r['status'] == 'ok']
    msg = f'{len(new_ok)}ファイル・{total}件取込みました' if new_ok else '新しいファイルはありませんでした'
    return jsonify({'status': 'ok', 'message': msg, 'results': results, 'imported': total})


# ─── 手動ファイルアップロード ────────────────────────────────────────────────

def _save_upload(file_obj):
    """アップロードされたファイルを一時保存してパスを返す"""
    ext = os.path.splitext(file_obj.filename)[1]
    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    file_obj.save(tmp_path)
    return tmp_path


@bp.route('/preview', methods=['POST'])
def preview_excel():
    """Excelの先頭数行を返してカラム確認に使う"""
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルなし'}), 400
    f = request.files['file']
    tmp = _save_upload(f)
    try:
        import openpyxl
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        ws = wb.active
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows.append(list(row))
            if i >= 5:
                break
        wb.close()
        return jsonify({'rows': rows})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        try: os.unlink(tmp)
        except: pass


@bp.route('/upload/daniel_purchase', methods=['POST'])
def upload_daniel_purchase():
    """手動アップロード: ダニエル発注ファイル (Watanabe_list_*.xlsm)"""
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'ファイルが選択されていません'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'status': 'error', 'message': 'ファイル名が不明です'}), 400
    tmp = _save_upload(f)
    try:
        from services.purchase_excel_parser import PurchaseExcelParser
        rows = PurchaseExcelParser().parse(tmp)
        imported = _upsert_purchases_from_excel(rows, agent='daniel')
        log = ImportLog.query.filter_by(filename=f.filename).first()
        if log:
            log.record_count = imported
            log.imported_at  = datetime.now()
        else:
            db.session.add(ImportLog(filename=f.filename, file_type='purchase_daniel',
                                     record_count=imported))
        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported, 'filename': f.filename,
                        'message': f'{imported}件取込みました'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        try: os.unlink(tmp)
        except: pass


@bp.route('/upload/daniel_ems', methods=['POST'])
def upload_daniel_ems():
    """手動アップロード: ダニエルEMSファイル (EMS발송리스트_*.xlsx)"""
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'ファイルが選択されていません'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'status': 'error', 'message': 'ファイル名が不明です'}), 400
    tmp = _save_upload(f)
    try:
        from services.ems_excel_parser import EmsExcelParser
        items = EmsExcelParser().parse(tmp)
        imported = _upsert_ems_items(items, agent='daniel')
        log = ImportLog.query.filter_by(filename=f.filename).first()
        if log:
            log.record_count = imported
            log.imported_at  = datetime.now()
        else:
            db.session.add(ImportLog(filename=f.filename, file_type='ems_daniel',
                                     record_count=imported))
        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported, 'filename': f.filename,
                        'message': f'{imported}件取込みました'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        try: os.unlink(tmp)
        except: pass


# ─── 全取込（cron + 手動ボタン共用）─────────────────────────────────────────

@bp.route('/all', methods=['POST'])
def import_all():
    results = {}

    # Cloudike ダニエル発注
    try:
        from services.cloudike_webdav import CloudikeWebDAV
        from services.purchase_excel_parser import PurchaseExcelParser
        webdav = CloudikeWebDAV()
        local_path, filename = webdav.download_latest_purchase()
        if local_path and filename and not ImportLog.query.filter_by(filename=filename).first():
            rows = PurchaseExcelParser().parse(local_path)
            imported = _upsert_purchases_from_excel(rows, agent='daniel')
            db.session.add(ImportLog(filename=filename, file_type='purchase_daniel', record_count=imported))
            db.session.commit()
            results['daniel_purchase'] = {'status': 'ok', 'imported': imported, 'filename': filename}
        else:
            results['daniel_purchase'] = {'status': 'ok', 'message': '新しい発注ファイルなし', 'imported': 0}
    except Exception as e:
        results['daniel_purchase'] = {'status': 'error', 'message': str(e)}

    # Cloudike ダニエルEMS
    try:
        from services.cloudike_webdav import CloudikeWebDAV
        from services.ems_excel_parser import EmsExcelParser
        webdav = CloudikeWebDAV()
        local_path, filename = webdav.download_latest_ems()
        if local_path and filename and not ImportLog.query.filter_by(filename=filename).first():
            items = EmsExcelParser().parse(local_path)
            imported = _upsert_ems_items(items, agent='daniel')
            db.session.add(ImportLog(filename=filename, file_type='ems_daniel', record_count=imported))
            db.session.commit()
            results['daniel_ems'] = {'status': 'ok', 'imported': imported, 'filename': filename}
        else:
            results['daniel_ems'] = {'status': 'ok', 'message': '新しいEMSファイルなし', 'imported': 0}
    except Exception as e:
        results['daniel_ems'] = {'status': 'error', 'message': str(e)}

    results['yahoo'] = {'status': 'ok', 'message': 'API設定後に有効化'}
    return jsonify({'status': 'ok', 'results': results})


# ─── 内部ヘルパー ─────────────────────────────────────────────────────────────

def _upsert_yahoo_order(order_raw: dict):
    """Yahoo受注1件を Order + OrderItem に upsert する。"""
    order_id_str = (order_raw.get('OrderId') or order_raw.get('order_id', '')).strip()
    if not order_id_str:
        return 0, 0, []

    ship_status  = str(order_raw.get('ShipStatus',  order_raw.get('ship_status',  '0')))
    order_status = str(order_raw.get('OrderStatus', order_raw.get('order_status', '')))

    # 顧客名
    buyer = (order_raw.get('BuyerInfo') or order_raw.get('Ship') or {})
    customer_name = (buyer.get('Name1', '') + ' ' + buyer.get('Name2', '')).strip() or order_raw.get('ShipName', '')

    # 受注日時（ISO8601 / YYYYMMDDHHmmss 両対応）
    order_time_str = str(order_raw.get('OrderTime', order_raw.get('order_time', '')))
    ordered_at = None
    for fmt in ('%Y-%m-%dT%H:%M:%S+09:00', '%Y-%m-%dT%H:%M:%S', '%Y%m%d%H%M%S'):
        try:
            ordered_at = datetime.strptime(order_time_str[:len(fmt)], fmt)
            break
        except Exception:
            continue
    if not ordered_at:
        try:
            ordered_at = datetime.fromisoformat(order_time_str.replace('+09:00', ''))
        except Exception:
            ordered_at = datetime.now()

    existing = Order.query.filter_by(yahoo_order_id=order_id_str).first()

    if ship_status in ('2', '3'):
        # 出荷済み・完了: Orderレコードのみ更新/作成（明細はスキップ）
        if existing:
            existing.yahoo_ship_status  = ship_status
            existing.yahoo_order_status = order_status
        else:
            db.session.add(Order(
                yahoo_order_id=order_id_str,
                customer_name=customer_name,
                yahoo_ship_status=ship_status,
                yahoo_order_status=order_status,
                ordered_at=ordered_at,
                status='shipped',
            ))
        return 0, 0, []

    new_order = 0
    if not existing:
        existing = Order(
            yahoo_order_id=order_id_str,
            customer_name=customer_name,
            yahoo_ship_status=ship_status,
            yahoo_order_status=order_status,
            ordered_at=ordered_at,
            status='pending',
        )
        db.session.add(existing)
        db.session.flush()
        new_order = 1
    else:
        existing.yahoo_ship_status  = ship_status
        existing.yahoo_order_status = order_status
        if customer_name:
            existing.customer_name = customer_name
        # ordered_atが壊れている（datetime.now()）場合のみ正しい値で上書き
        if ordered_at and abs((existing.ordered_at - ordered_at).days) > 1:
            existing.ordered_at = ordered_at

    # ── 明細 (OrderItem) ──
    items_raw = order_raw.get('ItemInfo', order_raw.get('Item', []))
    if isinstance(items_raw, dict):
        items_raw = [items_raw]

    new_items    = 0
    new_item_ids = []
    for item_raw in (items_raw or []):
        product_code = (item_raw.get('ItemId') or item_raw.get('item_id', '')).strip()
        if not product_code:
            continue
        existing_item = OrderItem.query.filter_by(
            order_id=existing.id, product_code=product_code
        ).first()
        if existing_item:
            continue
        try:
            qty = int(float(item_raw.get('Quantity', item_raw.get('quantity', 1))))
        except Exception:
            qty = 1
        oi = OrderItem(
            order_id=existing.id,
            product_code=product_code,
            product_name=item_raw.get('Title', item_raw.get('title', '')),
            quantity=qty,
            inventory_type='pending',
            status='pending',
        )
        db.session.add(oi)
        db.session.flush()
        new_item_ids.append(oi.id)
        new_items += 1

    return new_order, new_items, new_item_ids


def _upsert_purchases(rows, agent='tegu'):
    imported = 0
    for row in rows:
        product_code = row.get('商品コード', '')
        oi = OrderItem.query.filter_by(product_code=product_code).first()
        if not oi:
            continue
        if Purchase.query.filter_by(order_item_id=oi.id, product_code=product_code).first():
            continue
        try:
            ordered_at = datetime.strptime(row.get('発注日', ''), '%Y/%m/%d').date()
        except Exception:
            ordered_at = datetime.now().date()
        p = Purchase(
            order_item_id=oi.id, product_code=product_code,
            product_name=row.get('商品名', ''), quantity=int(row.get('数量', 1)),
            shop_name=row.get('発注先', ''), ordered_at=ordered_at,
            status='ordered', agent=agent,
        )
        db.session.add(p)
        imported += 1
    db.session.commit()
    return imported


def _upsert_purchases_from_excel(rows, agent='daniel'):
    """全行をupsert（purchase_no+product_codeで同定。statusは手動変更を保持）"""
    imported = 0
    updated = 0
    for row in rows:
        product_code = row.get('product_code', '')
        if not product_code:
            continue
        purchase_no = row.get('purchase_no', '') or ''
        order_id    = row.get('order_id', '') or ''

        # ── 既存チェック: purchase_no + product_code ──
        existing = None
        if purchase_no:
            existing = Purchase.query.filter_by(
                purchase_no=purchase_no, product_code=product_code
            ).first()

        if existing:
            # 更新（statusは保持、その他は上書き）
            existing.product_name = row.get('product_name') or existing.product_name
            existing.quantity     = row.get('quantity', existing.quantity) or existing.quantity
            existing.shop_name    = row.get('shop_name') or existing.shop_name
            existing.ordered_at   = row.get('ordered_at') or existing.ordered_at
            existing.memo         = row.get('memo') or existing.memo
            if order_id:
                existing.order_id = order_id
            existing.updated_at = datetime.now()
            updated += 1
        else:
            # OrderItemとの紐付けを試みる（なくても保存する）
            oi = OrderItem.query.filter_by(product_code=product_code).first()
            p = Purchase(
                purchase_no=purchase_no,
                order_id=order_id,
                order_item_id=oi.id if oi else None,
                product_code=product_code,
                product_name=row.get('product_name', ''),
                quantity=row.get('quantity', 1),
                shop_name=row.get('shop_name', ''),
                ordered_at=row.get('ordered_at') or datetime.now().date(),
                status='ordered',
                memo=row.get('memo', ''),
                agent=agent,
            )
            db.session.add(p)
            imported += 1

    db.session.commit()
    return imported + updated


def _upsert_ems_rows(rows, agent='tegu'):
    imported = 0
    for row in rows:
        ems_number = row.get('EMS番号', '')
        if not ems_number or Ems.query.filter_by(ems_number=ems_number).first():
            continue
        try:
            shipped_at = datetime.strptime(row.get('発送日', ''), '%Y/%m/%d').date()
        except Exception:
            shipped_at = datetime.now().date()
        ems = Ems(
            ems_number=ems_number, shipped_at=shipped_at,
            estimated_arrival=shipped_at + timedelta(days=3),
            status='in_transit', agent=agent,
        )
        db.session.add(ems)
        db.session.flush()
        product_code = row.get('商品コード', '')
        if product_code:
            oi = OrderItem.query.filter_by(product_code=product_code).first()
            if oi:
                db.session.add(EmsItem(
                    ems_id=ems.id, order_item_id=oi.id,
                    product_code=product_code, quantity=int(row.get('数量', 1)),
                ))
        imported += 1
    db.session.commit()
    return imported


def _upsert_ems_items(items, agent='daniel'):
    """全行をupsert（ems_numberで同定。arrived/statusは手動変更を保持）"""
    imported = 0
    updated = 0
    for item in items:
        ems_number = item.get('ems_number', '')
        if not ems_number:
            continue

        # shipped_at が None（発送日空欄）の場合は today() でなく None のままにする
        # → sort 汚染防止（今日付けになって一番上に来てしまうバグを防ぐ）
        shipped_at    = item.get('shipped_at')  # None OK
        product_code  = item.get('product_code', '')
        purchase_date = item.get('purchase_date', '') or ''   # B列: 구매日 作成日
        purchase_no   = item.get('purchase_no', '') or ''     # F列: 구매番号
        order_id      = item.get('order_id', '') or ''

        # ── Ems のUPSERT ──
        ems = Ems.query.filter_by(ems_number=ems_number).first()
        if not ems:
            ems = Ems(
                ems_number=ems_number,
                purchase_no=purchase_no,
                order_id=order_id,
                shipped_at=shipped_at,   # None の場合もそのままセット
                estimated_arrival=(shipped_at + timedelta(days=3)) if shipped_at else None,
                status='in_transit',
                agent=agent,
            )
            db.session.add(ems)
            db.session.flush()
        else:
            # 発送日・発注NOは更新（statusはarrived保持）
            if shipped_at:  # None で上書きしない
                ems.shipped_at = shipped_at
            if purchase_no:
                ems.purchase_no = purchase_no
            if order_id:
                ems.order_id = order_id

        # ── EmsItem のUPSERT ──
        # キー: ems_id + product_code + purchase_no（purchase_no がある場合は精確に区別）
        if product_code:
            if purchase_no:
                existing_item = EmsItem.query.filter_by(
                    ems_id=ems.id, product_code=product_code, purchase_no=purchase_no
                ).first()
                # purchase_no なし（旧データ）で同一product_codeが存在する場合はそちらを更新
                if not existing_item:
                    existing_item = EmsItem.query.filter(
                        EmsItem.ems_id == ems.id,
                        EmsItem.product_code == product_code,
                        EmsItem.purchase_no == None,
                    ).first()
            else:
                existing_item = EmsItem.query.filter_by(
                    ems_id=ems.id, product_code=product_code
                ).filter(EmsItem.purchase_no == None).first()

            if existing_item:
                existing_item.quantity     = item.get('quantity', existing_item.quantity)
                existing_item.product_name = item.get('product_name') or existing_item.product_name
                existing_item.purchase_date = purchase_date or existing_item.purchase_date
                if purchase_no:
                    existing_item.purchase_no = purchase_no
                updated += 1
            else:
                # ① purchase_no → Purchase → order_item_id の順で紐付けを試みる
                oi = None
                if purchase_no:
                    from models.purchase import Purchase
                    p = Purchase.query.filter_by(
                        purchase_no=purchase_no, product_code=product_code
                    ).first()
                    if p and p.order_item_id:
                        from models.order_item import OrderItem as OI
                        oi = OI.query.get(p.order_item_id)
                # ② product_code フォールバック
                if oi is None:
                    oi = OrderItem.query.filter_by(product_code=product_code).first()

                db.session.add(EmsItem(
                    ems_id=ems.id,
                    order_item_id=oi.id if oi else None,
                    purchase_date=purchase_date or None,
                    purchase_no=purchase_no or None,
                    product_code=product_code,
                    product_name=item.get('product_name', ''),
                    quantity=item.get('quantity', 1),
                ))
                imported += 1

    db.session.commit()
    return imported + updated


def _upsert_import_log(filename, file_type, record_count):
    """ImportLog を upsert（新規 or 更新）"""
    log = ImportLog.query.filter_by(filename=filename).first()
    if log:
        log.record_count = record_count
        log.imported_at  = datetime.now()
    else:
        db.session.add(ImportLog(filename=filename, file_type=file_type, record_count=record_count))


def run_all_imports_job(app):
    """APScheduler から呼び出す用（1時間ごと自動実行）
    ・Cloudike の全ファイルをupsert取込
    ・ローカルフォルダ（Google Drive）も同様
    """
    with app.app_context():
        # ─── Cloudike WebDAV（全ファイルをupsert）──────────────────────────
        try:
            from services.cloudike_webdav import CloudikeWebDAV
            from services.purchase_excel_parser import PurchaseExcelParser
            from services.ems_excel_parser import EmsExcelParser
            webdav = CloudikeWebDAV()

            # 発注ファイル：全件取込
            for file_info in webdav.list_all_purchase_files():
                fname = file_info['name']
                try:
                    local_path = webdav.download_file(file_info['href'], fname)
                    rows = PurchaseExcelParser().parse(local_path)
                    cnt = _upsert_purchases_from_excel(rows, agent='daniel')
                    _upsert_import_log(fname, 'purchase_daniel', cnt)
                    db.session.commit()
                    app.logger.info(f'Auto import purchase: {fname} → {cnt} 件')
                except Exception as e:
                    db.session.rollback()
                    app.logger.error(f'Auto import purchase error [{fname}]: {e}')

            # EMSファイル：全件取込
            for file_info in webdav.list_all_ems_files():
                fname = file_info['name']
                try:
                    local_path = webdav.download_file(file_info['href'], fname)
                    items = EmsExcelParser().parse(local_path)
                    cnt = _upsert_ems_items(items, agent='daniel')
                    _upsert_import_log(fname, 'ems_daniel', cnt)
                    db.session.commit()
                    app.logger.info(f'Auto import EMS: {fname} → {cnt} 件')
                except Exception as e:
                    db.session.rollback()
                    app.logger.error(f'Auto import EMS error [{fname}]: {e}')

        except Exception as e:
            app.logger.error(f'Auto import WebDAV error: {e}')

        # ─── フォルダ監視（Google Drive / ローカル）────────────────────────
        try:
            from services.purchase_excel_parser import PurchaseExcelParser
            from services.ems_excel_parser import EmsExcelParser
            p_folder = os.getenv('WATCH_FOLDER_PURCHASE', '')
            for fname, fpath in _scan_folder(p_folder, PURCHASE_PATTERN_RE):
                try:
                    rows = PurchaseExcelParser().parse(fpath)
                    cnt = _upsert_purchases_from_excel(rows, agent='daniel')
                    _upsert_import_log(fname, 'purchase_daniel', cnt)
                    db.session.commit()
                    app.logger.info(f'Folder import purchase: {fname} → {cnt} 件')
                except Exception as e:
                    db.session.rollback()
                    app.logger.error(f'Folder import purchase error [{fname}]: {e}')

            e_folder = os.getenv('WATCH_FOLDER_EMS', '')
            for fname, fpath in _scan_folder(e_folder, EMS_PATTERN_RE):
                try:
                    items = EmsExcelParser().parse(fpath)
                    cnt = _upsert_ems_items(items, agent='daniel')
                    _upsert_import_log(fname, 'ems_daniel', cnt)
                    db.session.commit()
                    app.logger.info(f'Folder import EMS: {fname} → {cnt} 件')
                except Exception as e:
                    db.session.rollback()
                    app.logger.error(f'Folder import EMS error [{fname}]: {e}')
        except Exception as e:
            app.logger.error(f'Auto import folder error: {e}')
