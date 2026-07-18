"""データ取込ルート"""
import os, tempfile
from flask import Blueprint, request, jsonify, current_app
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.inventory import Inventory
from models.alert import Alert
from datetime import datetime

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

        # ── 在庫台帳: 注文出庫・キャンセル戻しを記録 ──────────────────
        ledger_out = 0
        ledger_ret = {'returned': 0, 'alerted': 0}
        try:
            from services.stock_ledger import apply_order_out, sync_cancel_returns
            ledger_out = apply_order_out(new_item_ids)
            ledger_ret = sync_cancel_returns()
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            current_app.logger.error(f'Ledger hook error: {e}')

        return jsonify({
            'status':          'ok',
            'imported_orders': imported_orders,
            'imported_items':  imported_items,
            'alloc':           alloc_stats,
            'ledger_out':      ledger_out,
            'ledger_returns':  ledger_ret,
            'message': (f'{imported_orders}件の受注・{imported_items}件の明細を取込みました'
                        f' | 引当: 即納{alloc_stats["fully_allocated"]}件'
                        f' / 部分{alloc_stats["partial"]}件'
                        f' / お取り寄せ{alloc_stats["tori"]}件'
                        f' / 不足{alloc_stats["shortage"]}件'
                        f' | 台帳: 出庫{ledger_out}件 / 戻し{ledger_ret["returned"]}件'),
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


def run_yahoo_stock_full_import():
    """Yahoo在庫フル同期の本体。`/import/yahoo_stock` と CLI から利用する。

    Returns:
        dict: HTTP の JSON と同形の結果（status, message, 件数, alloc など）

    Raises:
        Exception: DB / API 等の失敗時（呼び出し側で rollback / exit code 処理）
    """
    from services.yahoo_api import YahooAPI
    api = YahooAPI()

    # ① 在庫CSV取得（type=2）
    stock_items = api.download_inventory_csv()
    if not stock_items:
        return {'status': 'ok', 'message': '在庫データがありませんでした', 'updated': 0}

    # ② 商品CSV取得（type=1）→ code をキーにした辞書を作成
    try:
        item_list = api.download_item_csv()
        item_names = {i['code']: i for i in item_list}
    except Exception as e:
        current_app.logger.warning(f'商品CSV取得失敗（在庫のみ更新）: {e}')
        item_names = {}

    imported = 0
    updated = 0
    now = datetime.now()

    for item in stock_items:
        code = item['code']
        sub_code = item['sub_code'] or None
        qty = item['quantity']
        if not code or qty is None:
            continue  # コードなし or 在庫無限大はスキップ

        # サブコード末尾が 'b' → お取り寄せ、それ以外 → 即納
        is_immediate = _is_immediate(sub_code)

        # 商品名・価格（商品CSV から補完）
        item_info = item_names.get(code, {})
        product_name = item_info.get('name', '')
        price = item_info.get('price', 0)

        existing = Inventory.query.filter_by(
            product_code=code,
            product_sub_code=sub_code,
            inventory_type='yahoo',
        ).first()

        if existing:
            existing.yahoo_stock = qty
            existing.quantity = qty
            existing.is_immediate = is_immediate
            existing.available_qty = max(0, qty - (existing.reserved_qty or 0))
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
    return {
        'status': 'ok',
        'imported': imported,
        'updated': updated,
        'total': total,
        'immediate': immediate_cnt,
        'alloc': alloc_stats,
        'message': (f'Yahoo在庫同期完了: {imported}件新規 / {updated}件更新（計{total}件）'
                      f' | 即納:{immediate_cnt}件'
                      f' | 引当: 即納{alloc_stats["fully_allocated"]}件'
                      f' / 部分{alloc_stats["partial"]}件'
                      f' / お取り寄せ{alloc_stats["tori"]}件'
                      f' / 不足{alloc_stats["shortage"]}件'),
    }


@bp.route('/yahoo_stock', methods=['POST'])
def import_yahoo_stock():
    """【初回・フル同期】Yahoo在庫＋商品名を一括取得して Inventory を差分更新。
    - 在庫CSV (type=2): quantity, allow_overdraft
    - 商品CSV (type=1): 商品名, 価格
    - allow_overdraft=0 かつ quantity>0 → is_immediate=True（即納）
    - allow_overdraft=1 → is_immediate=False（お取り寄せ）
    """
    try:
        result = run_yahoo_stock_full_import()
        return jsonify(result)
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


# ─── 手動ファイルアップロード ────────────────────────────────────────────────

def _save_upload(file_obj):
    """アップロードされたファイルを一時保存してパスを返す"""
    ext = os.path.splitext(file_obj.filename)[1]
    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    file_obj.save(tmp_path)
    return tmp_path


# ─── 全取込（cron + 手動ボタン共用）─────────────────────────────────────────

@bp.route('/all', methods=['POST'])
def import_all():
    results = {}

    results['yahoo'] = {'status': 'ok', 'message': 'API設定後に有効化'}

    # ─── 台帳整合性チェック ────────────────────────────────────────
    try:
        from services.stock_ledger import verify_cache_integrity
        mismatches = verify_cache_integrity()
        db.session.commit()
    except Exception:
        db.session.rollback()
        mismatches = []

    return jsonify({'status': 'ok', 'results': results, 'ledger_mismatch_fixed': len(mismatches)})


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
            # 初見で既に出荷済み: 台帳出庫が記録されないためアラート(取込停止期間の検知用)
            if ordered_at and (datetime.now() - ordered_at).days <= 35:
                db.session.add(Alert(
                    alert_type='ledger_out_missing',
                    product_code=None,
                    message=(f'注文 {order_id_str} は初回取込時点で既に出荷済みのため、'
                             f'台帳の出庫記録がありません。必要なら手動出庫してください。'),
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


def run_all_imports_job(app):
    """APScheduler / systemdタイマーから呼び出す定期ジョブ。
    ダニエル・テグ(大邱)の取込廃止後は台帳整合性チェックのみ実施する。
    """
    with app.app_context():
        # ─── 台帳整合性チェック ────────────────────────────────────────
        try:
            from services.stock_ledger import verify_cache_integrity
            mismatches = verify_cache_integrity()
            db.session.commit()
            if mismatches:
                app.logger.warning(f'Ledger integrity: {len(mismatches)} 件修正 {mismatches}')
        except Exception as e:
            db.session.rollback()
            app.logger.error(f'Ledger integrity check error: {e}')
