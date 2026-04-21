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

AGENT_LABELS = {'daniel': 'ダニエル', 'tegu': 'テグ'}


PER_PAGE = 200


def _get_purchase_data(agent, q='', status_filter='', page=1):
    query = db.session.query(Purchase, OrderItem, Order).outerjoin(
        OrderItem, Purchase.order_item_id == OrderItem.id
    ).outerjoin(
        Order, OrderItem.order_id == Order.id
    ).filter(Purchase.agent == agent)

    if q:
        query = query.filter(
            db.or_(
                Purchase.product_code.contains(q),
                Purchase.product_name.contains(q),
                Purchase.shop_name.contains(q),
                Purchase.purchase_no.contains(q),
                Purchase.order_id.contains(q),
                Order.yahoo_order_id.contains(q),
            )
        )
    if status_filter:
        query = query.filter(Purchase.status == status_filter)

    total = query.count()
    rows  = query.order_by(Purchase.ordered_at.desc()) \
                 .offset((page - 1) * PER_PAGE).limit(PER_PAGE).all()

    purchase_data = [
        {'purchase': p, 'purchase_no': p.purchase_no or '', 'order_item': oi, 'order': o}
        for p, oi, o in rows
    ]
    return purchase_data, total


@bp.route('/')
def index():
    agent         = request.args.get('agent', 'daniel')
    q             = request.args.get('q', '').strip()
    status_filter = request.args.get('status', '')
    page          = max(1, int(request.args.get('page', 1)))

    purchase_data, total = _get_purchase_data(agent, q, status_filter, page)
    total_pages = max(1, (total + PER_PAGE - 1) // PER_PAGE)

    return render_template('purchases.html',
                           purchase_data=purchase_data,
                           agent=agent,
                           agent_label=AGENT_LABELS.get(agent, agent),
                           q=q, status_filter=status_filter,
                           page=page, total_pages=total_pages, total=total)


@bp.route('/update/<int:purchase_id>', methods=['POST'])
def update(purchase_id):
    p = Purchase.query.get_or_404(purchase_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    allowed = {'product_code', 'product_name', 'quantity', 'shop_name',
               'ordered_at', 'status', 'memo'}
    if field not in allowed:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        if field == 'quantity':
            setattr(p, field, int(value))
        elif field == 'ordered_at':
            setattr(p, field, datetime.strptime(value, '%Y-%m-%d').date() if value else None)
        else:
            setattr(p, field, value)
        p.updated_at = datetime.now()
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


@bp.route('/allocate', methods=['POST'])
def allocate():
    purchase_ids = request.json.get('purchase_ids', [])
    results = []
    for pid in purchase_ids:
        p = Purchase.query.get(pid)
        if not p:
            continue
        # order_item_id が無い場合は product_code で検索
        item = None
        if p.order_item_id:
            item = OrderItem.query.get(p.order_item_id)
        if not item and p.product_code:
            item = OrderItem.query.filter_by(product_code=p.product_code).first()
        if not item:
            results.append({'purchase_id': pid, 'status': 'no_order_item'})
            continue
        if item.status in ('fully_allocated', 'shipped'):
            results.append({'purchase_id': pid, 'status': 'already_done'})
            continue
        if not Allocation.query.filter_by(order_item_id=item.id, allocation_type='仮引当').first():
            db.session.add(Allocation(
                order_item_id=item.id, inventory_type='お取り寄せ',
                allocation_type='仮引当', quantity=p.quantity,
            ))
            item.status = 'provisional_allocated'
            results.append({'purchase_id': pid, 'status': 'provisional_allocated'})
        else:
            results.append({'purchase_id': pid, 'status': 'already_provisional'})
    db.session.commit()
    return jsonify({'results': results})


@bp.route('/csv')
def export_csv():
    agent = request.args.get('agent', 'daniel')
    purchase_data = _get_purchase_data(agent)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['発注NO', '受注番号', '商品コード', '商品名', '数量', '発注先', '発注日', 'ステータス', 'メモ'])
    for d in purchase_data:
        p = d['purchase']
        o = d['order']
        writer.writerow([
            p.purchase_no or '',
            o.yahoo_order_id if o else (p.order_id or ''),
            p.product_code,
            p.product_name or '',
            p.quantity,
            p.shop_name or '',
            p.ordered_at.strftime('%Y/%m/%d') if p.ordered_at else '',
            p.status_label,
            p.memo or '',
        ])
    output.seek(0)
    label = AGENT_LABELS.get(agent, agent)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename={label}_purchases_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
