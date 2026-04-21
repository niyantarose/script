import csv
import io
from flask import Blueprint, render_template, request, jsonify, Response
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.allocation import Allocation
from models.inventory import Inventory
from models.ems_item import EmsItem
from datetime import datetime
from sqlalchemy import or_

bp = Blueprint('orders', __name__, url_prefix='/orders')


@bp.route('/')
def index():
    q = request.args.get('q', '').strip()
    status_filter = request.args.get('status', '')
    # 商品コードで絞る場合は order_items から引く
    if q:
        item_order_ids = [
            oi.order_id for oi in
            OrderItem.query.filter(OrderItem.product_code.contains(q)).all()
        ]
        query = Order.query.filter(
            or_(
                Order.yahoo_order_id.contains(q),
                Order.customer_name.contains(q),
                Order.id.in_(item_order_ids),
            )
        )
    else:
        query = Order.query
    if status_filter:
        item_order_ids = [
            oi.order_id for oi in
            OrderItem.query.filter_by(status=status_filter).all()
        ]
        query = query.filter(Order.id.in_(item_order_ids))
    orders = query.order_by(Order.ordered_at.desc()).all()

    # order単位でグルーピング（注文ごとに枠線を出すため）
    grouped = []
    for o in orders:
        items = OrderItem.query.filter_by(order_id=o.id).all()
        # EMS番号を収集
        ems_nums = set()
        for item in items:
            for ei in EmsItem.query.filter_by(order_item_id=item.id).all():
                ems_nums.add(ei.ems.ems_number)
        grouped.append({
            'order': o,
            'order_items': items,
            'ems_numbers': '、'.join(sorted(ems_nums)) if ems_nums else '',
        })

    return render_template('orders.html', grouped=grouped, q=q, status_filter=status_filter)


@bp.route('/allocate', methods=['POST'])
def allocate():
    """引当実行（受注リストから）"""
    item_ids = request.json.get('item_ids', [])
    results = []

    for item_id in item_ids:
        item = OrderItem.query.get(item_id)
        if not item:
            continue

        needed = item.quantity - item.allocated_qty
        if needed <= 0:
            results.append({'item_id': item_id, 'status': 'already_allocated'})
            continue

        # 即納在庫チェック
        inv = Inventory.query.filter_by(
            product_code=item.product_code,
            inventory_type='即納'
        ).first()

        if inv and inv.available_qty >= needed:
            # 全量即納引当
            inv.reserved_qty += needed
            inv.available_qty -= needed
            item.allocated_qty = item.quantity
            item.status = 'fully_allocated'

            alloc = Allocation(
                order_item_id=item.id,
                inventory_type='即納',
                allocation_type='本引当',
                quantity=needed,
            )
            db.session.add(alloc)
            results.append({'item_id': item_id, 'status': 'fully_allocated'})

        elif inv and inv.available_qty > 0:
            # 部分引当
            alloc_qty = inv.available_qty
            inv.reserved_qty += alloc_qty
            inv.available_qty = 0
            item.allocated_qty += alloc_qty
            item.status = 'partial_waiting'

            alloc = Allocation(
                order_item_id=item.id,
                inventory_type='即納',
                allocation_type='本引当',
                quantity=alloc_qty,
            )
            db.session.add(alloc)
            results.append({'item_id': item_id, 'status': 'partial_waiting'})

        elif item.inventory_type == 'お取り寄せ':
            item.status = 'provisional_allocated'
            results.append({'item_id': item_id, 'status': 'provisional_allocated'})

        else:
            item.status = 'shortage'
            results.append({'item_id': item_id, 'status': 'shortage'})

    db.session.commit()
    return jsonify({'results': results})


@bp.route('/update/order/<int:order_id>', methods=['POST'])
def update_order(order_id):
    o = Order.query.get_or_404(order_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    allowed = {'customer_name', 'yahoo_ship_status', 'status', 'delay_memo', 'desired_delivery_date'}
    if field not in allowed:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        if field == 'desired_delivery_date':
            o.desired_delivery_date = datetime.strptime(value, '%Y-%m-%d').date() if value else None
        else:
            setattr(o, field, value)
        o.updated_at = datetime.now()
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


@bp.route('/update/item/<int:item_id>', methods=['POST'])
def update_item(item_id):
    item = OrderItem.query.get_or_404(item_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    allowed = {'product_code', 'quantity', 'status', 'allocated_qty', 'inventory_type'}
    if field not in allowed:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        if field in ('quantity', 'allocated_qty'):
            setattr(item, field, int(value))
        else:
            setattr(item, field, value)
        item.updated_at = datetime.now()
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


@bp.route('/csv')
def export_csv():
    orders = Order.query.order_by(Order.ordered_at.desc()).all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['注文番号', '注文日', 'お客様名', '商品コード', '商品名', '数量',
                     '在庫種別', '引当ステータス', '引当数', '経過日数'])

    for o in orders:
        items = OrderItem.query.filter_by(order_id=o.id).all()
        for item in items:
            writer.writerow([
                o.yahoo_order_id,
                o.ordered_at.strftime('%Y/%m/%d'),
                o.customer_name or '',
                item.product_code,
                item.product_name or '',
                item.quantity,
                item.inventory_type,
                item.status_label,
                item.allocated_qty,
                o.days_elapsed,
            ])

    output.seek(0)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=orders_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
