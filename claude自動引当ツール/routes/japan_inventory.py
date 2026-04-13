import csv
import io
from flask import Blueprint, render_template, request, jsonify, Response
from models import db
from models.japan_inventory import JapanInventoryStaging
from models.ems import Ems
from models.ems_item import EmsItem
from models.order_item import OrderItem
from models.order import Order
from models.allocation import Allocation
from models.inventory import Inventory
from datetime import datetime

bp = Blueprint('japan_inventory', __name__, url_prefix='/japan')


@bp.route('/')
def index():
    staging = db.session.query(
        JapanInventoryStaging, EmsItem, Ems
    ).join(
        EmsItem, JapanInventoryStaging.ems_item_id == EmsItem.id
    ).join(
        Ems, EmsItem.ems_id == Ems.id
    ).order_by(Ems.arrived_at.desc()).all()

    staging_data = []
    for jis, ei, e in staging:
        assigned_order = None
        if jis.assigned_order_item_id:
            oi = OrderItem.query.get(jis.assigned_order_item_id)
            if oi:
                assigned_order = Order.query.get(oi.order_id)

        staging_data.append({
            'staging': jis,
            'ems_item': ei,
            'ems': e,
            'assigned_order': assigned_order,
        })

    return render_template('japan_inventory.html', staging_data=staging_data)


@bp.route('/update_status', methods=['POST'])
def update_status():
    """仕分けステータス更新"""
    staging_id = request.json.get('staging_id')
    new_status = request.json.get('status')
    reason = request.json.get('reason', '')

    jis = JapanInventoryStaging.query.get(staging_id)
    if not jis:
        return jsonify({'error': 'Not found'}), 404

    if new_status == 'to_japan_stock':
        jis.status = 'to_japan_stock'
    elif new_status == 'excluded':
        jis.status = 'excluded'
        jis.excluded_reason = reason
    elif new_status == 'returned_to_ems':
        jis.status = 'returned_to_ems'
    elif new_status == 'waiting':
        jis.status = 'waiting'
        jis.assigned_order_item_id = None

    db.session.commit()
    return jsonify({'status': 'ok', 'new_status': jis.status, 'new_status_label': jis.status_label})


@bp.route('/assign_order', methods=['POST'])
def assign_order():
    """受注に手動引当"""
    staging_id = request.json.get('staging_id')
    order_item_id = request.json.get('order_item_id')

    jis = JapanInventoryStaging.query.get(staging_id)
    if not jis:
        return jsonify({'error': 'Not found'}), 404

    item = OrderItem.query.get(order_item_id)
    if not item:
        return jsonify({'error': 'Order item not found'}), 404

    jis.status = 'assigned_to_order'
    jis.assigned_order_item_id = order_item_id

    # 引当記録を作成
    alloc = Allocation(
        order_item_id=order_item_id,
        inventory_type='お取り寄せ',
        allocation_type='手動',
        quantity=jis.quantity,
        ems_item_id=jis.ems_item_id,
        allocated_by='staff',
    )
    db.session.add(alloc)

    # 受注明細の引当数を更新
    item.allocated_qty += jis.quantity
    if item.allocated_qty >= item.quantity:
        item.status = 'fully_allocated'
    else:
        item.status = 'partial_waiting'

    db.session.commit()

    order = Order.query.get(item.order_id)
    return jsonify({
        'status': 'ok',
        'assigned_order_id': order.yahoo_order_id if order else '',
    })


@bp.route('/reflect', methods=['POST'])
def reflect_to_yahoo():
    """日本在庫をYahoo在庫に反映"""
    targets = JapanInventoryStaging.query.filter_by(status='to_japan_stock').all()
    reflected_count = 0

    for jis in targets:
        # 在庫テーブルを更新（実際のYahoo API呼び出しは別途実装）
        inv = Inventory.query.filter_by(
            product_code=jis.product_code,
            inventory_type='即納'
        ).first()

        if inv:
            inv.quantity += jis.quantity
            inv.available_qty += jis.quantity
        else:
            inv = Inventory(
                product_code=jis.product_code,
                product_sub_code=jis.product_sub_code,
                inventory_type='即納',
                quantity=jis.quantity,
                reserved_qty=0,
                available_qty=jis.quantity,
            )
            db.session.add(inv)

        jis.status = 'reflected'
        jis.reflected_at = datetime.now()
        reflected_count += 1

    db.session.commit()
    return jsonify({'status': 'ok', 'reflected_count': reflected_count})


@bp.route('/csv')
def export_csv():
    staging = db.session.query(
        JapanInventoryStaging, EmsItem, Ems
    ).join(
        EmsItem, JapanInventoryStaging.ems_item_id == EmsItem.id
    ).join(
        Ems, EmsItem.ems_id == Ems.id
    ).order_by(Ems.arrived_at.desc()).all()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['EMS番号', '入荷日', '商品コード', '数量', '仕分けステータス', '引当先受注番号'])

    for jis, ei, e in staging:
        assigned_order_id = ''
        if jis.assigned_order_item_id:
            oi = OrderItem.query.get(jis.assigned_order_item_id)
            if oi:
                o = Order.query.get(oi.order_id)
                assigned_order_id = o.yahoo_order_id if o else ''

        writer.writerow([
            e.ems_number,
            e.arrived_at.strftime('%Y/%m/%d') if e.arrived_at else '',
            jis.product_code,
            jis.quantity,
            jis.status_label,
            assigned_order_id,
        ])

    output.seek(0)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=japan_inventory_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
