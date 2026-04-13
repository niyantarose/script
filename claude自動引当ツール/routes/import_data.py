"""データ取込ルート"""
from flask import Blueprint, request, jsonify
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.ems import Ems
from models.ems_item import EmsItem
from models.inventory import Inventory
from datetime import datetime, timedelta

bp = Blueprint('import_data', __name__, url_prefix='/import')


@bp.route('/yahoo_orders', methods=['POST'])
def import_yahoo_orders():
    """Yahoo APIから受注を取込"""
    try:
        from services.yahoo_api import YahooAPI
        api = YahooAPI()
        seller_id = request.json.get('seller_id', '')
        # TODO: APIから受注データを取得してDBに保存
        # 実際のAPI呼び出しはYahoo API設定後に有効化
        return jsonify({'status': 'ok', 'message': 'Yahoo受注取込はAPI設定後に有効化されます'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/yahoo_stock', methods=['POST'])
def import_yahoo_stock():
    """Yahoo APIから在庫を取込"""
    try:
        from services.yahoo_api import YahooAPI
        api = YahooAPI()
        return jsonify({'status': 'ok', 'message': 'Yahoo在庫取込はAPI設定後に有効化されます'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/google_purchases', methods=['POST'])
def import_google_purchases():
    """Google Sheetsから発注リストを取込"""
    try:
        from services.google_sheets import GoogleSheetsAPI
        api = GoogleSheetsAPI()
        rows = api.fetch_purchases()

        imported = 0
        for row in rows:
            # 受注明細を探す
            order_item = OrderItem.query.filter_by(
                product_code=row.get('商品コード', '')
            ).first()
            if not order_item:
                continue

            # 重複チェック
            existing = Purchase.query.filter_by(
                order_item_id=order_item.id,
                product_code=row.get('商品コード', ''),
            ).first()
            if existing:
                continue

            p = Purchase(
                order_item_id=order_item.id,
                product_code=row.get('商品コード', ''),
                product_name=row.get('商品名', ''),
                quantity=int(row.get('数量', 1)),
                shop_name=row.get('発注先', ''),
                ordered_at=datetime.strptime(row.get('発注日', ''), '%Y/%m/%d').date() if row.get('発注日') else datetime.now().date(),
                status='ordered',
            )
            db.session.add(p)
            imported += 1

        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/google_ems', methods=['POST'])
def import_google_ems():
    """Google SheetsからEMSリストを取込"""
    try:
        from services.google_sheets import GoogleSheetsAPI
        api = GoogleSheetsAPI()
        rows = api.fetch_ems_list()

        imported = 0
        for row in rows:
            ems_number = row.get('EMS番号', '')
            if not ems_number:
                continue

            # 重複チェック
            existing = Ems.query.filter_by(ems_number=ems_number).first()
            if existing:
                continue

            shipped_str = row.get('発送日', '')
            shipped_at = datetime.strptime(shipped_str, '%Y/%m/%d').date() if shipped_str else datetime.now().date()

            ems = Ems(
                ems_number=ems_number,
                shipped_at=shipped_at,
                estimated_arrival=shipped_at + timedelta(days=2),
                status='in_transit',
            )
            db.session.add(ems)
            db.session.flush()

            # EMS明細もシートに含まれていればパース
            product_code = row.get('商品コード', '')
            if product_code:
                order_item = OrderItem.query.filter_by(product_code=product_code).first()
                if order_item:
                    ei = EmsItem(
                        ems_id=ems.id,
                        order_item_id=order_item.id,
                        product_code=product_code,
                        quantity=int(row.get('数量', 1)),
                    )
                    db.session.add(ei)

            imported += 1

        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500


@bp.route('/cloudike_ems', methods=['POST'])
def import_cloudike_ems():
    """Cloudike WebDAVからEMSファイルを取込"""
    try:
        from services.cloudike_webdav import CloudikeWebDAV
        from services.ems_excel_parser import EmsExcelParser

        webdav = CloudikeWebDAV()
        file_path = webdav.download_latest_ems()
        if not file_path:
            return jsonify({'status': 'ok', 'message': '新しいEMSファイルがありません', 'imported': 0})

        parser = EmsExcelParser()
        items = parser.parse(file_path)

        imported = 0
        for item in items:
            ems_number = item.get('ems_number', '')
            if not ems_number:
                continue

            # EMS便を取得または作成
            ems = Ems.query.filter_by(ems_number=ems_number).first()
            if not ems:
                shipped_at = item.get('shipped_at', datetime.now().date())
                ems = Ems(
                    ems_number=ems_number,
                    shipped_at=shipped_at,
                    estimated_arrival=shipped_at + timedelta(days=2) if hasattr(shipped_at, 'day') else datetime.now().date(),
                    status='in_transit',
                )
                db.session.add(ems)
                db.session.flush()

            # 受注明細を探して紐付け
            product_code = item.get('product_code', '')
            order_item = OrderItem.query.filter_by(product_code=product_code).first()
            if order_item:
                existing = EmsItem.query.filter_by(
                    ems_id=ems.id, order_item_id=order_item.id
                ).first()
                if not existing:
                    ei = EmsItem(
                        ems_id=ems.id,
                        order_item_id=order_item.id,
                        product_code=product_code,
                        quantity=item.get('quantity', 1),
                    )
                    db.session.add(ei)
                    imported += 1

        db.session.commit()
        return jsonify({'status': 'ok', 'imported': imported})
    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500
