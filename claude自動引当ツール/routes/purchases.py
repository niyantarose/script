import csv
import io
from flask import Blueprint, render_template, request, jsonify, Response
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.allocation import Allocation
from datetime import datetime

bp = Blueprint('purchases', __name__, url_prefix='/purchases')


@bp.route('/')
def index():
    purchases = db.session.query(Purchase, OrderItem, Order).join(
        OrderItem, Purchase.order_item_id == OrderItem.id
    ).join(
        Order, OrderItem.order_id == Order.id
    ).order_by(Purchase.ordered_at.desc()).all()

    # 未発注チェック（発注漏れ）
    missing = db.session.query(OrderItem, Order).join(
        Order, OrderItem.order_id == Order.id
    ).filter(
        OrderItem.inventory_type == 'お取り寄せ',
        ~OrderItem.id.in_(db.session.query(Purchase.order_item_id).distinct()),
        OrderItem.status.notin_(['shipped', 'fully_allocated'])
    ).all()

    purchase_no_counter = {}
    purchase_data = []
    for p, oi, o in purchases:
        # 発注番号を生成（shop_name + 日付ベース）
        key = f"Wata{p.ordered_at.strftime('%y%m%d')}" if p.ordered_at else 'N/A'
        if key not in purchase_no_counter:
            purchase_no_counter[key] = 0
        purchase_no_counter[key] += 1
        purchase_no = f"{key}_{purchase_no_counter[key]:02d}"

        purchase_data.append({
            'purchase': p,
            'purchase_no': purchase_no,
            'order_item': oi,
            'order': o,
        })

    return render_template('purchases.html',
                           purchase_data=purchase_data,
                           missing_count=len(missing))


@bp.route('/allocate', methods=['POST'])
def allocate():
    """発注リストからの引当実行（仮引当）"""
    purchase_ids = request.json.get('purchase_ids', [])
    results = []

    for pid in purchase_ids:
        p = Purchase.query.get(pid)
        if not p:
            continue

        item = OrderItem.query.get(p.order_item_id)
        if not item:
            continue

        if item.status in ('fully_allocated', 'shipped'):
            results.append({'purchase_id': pid, 'status': 'already_done'})
            continue

        # 仮引当を記録
        existing = Allocation.query.filter_by(
            order_item_id=item.id,
            allocation_type='仮引当'
        ).first()

        if not existing:
            alloc = Allocation(
                order_item_id=item.id,
                inventory_type='お取り寄せ',
                allocation_type='仮引当',
                quantity=p.quantity,
            )
            db.session.add(alloc)
            item.status = 'provisional_allocated'
            results.append({'purchase_id': pid, 'status': 'provisional_allocated'})
        else:
            results.append({'purchase_id': pid, 'status': 'already_provisional'})

    db.session.commit()
    return jsonify({'results': results})


@bp.route('/csv')
def export_csv():
    purchases = db.session.query(Purchase, OrderItem, Order).join(
        OrderItem, Purchase.order_item_id == OrderItem.id
    ).join(
        Order, OrderItem.order_id == Order.id
    ).order_by(Purchase.ordered_at.desc()).all()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['受注番号', '商品コード', '商品名', '数量', '発注先', '発注日', 'ステータス'])

    for p, oi, o in purchases:
        writer.writerow([
            o.yahoo_order_id,
            p.product_code,
            p.product_name or oi.product_name or '',
            p.quantity,
            p.shop_name or '',
            p.ordered_at.strftime('%Y/%m/%d') if p.ordered_at else '',
            p.status_label,
        ])

    output.seek(0)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=purchases_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
