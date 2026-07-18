import uuid

from flask import Blueprint, jsonify, render_template, request

from models import db, Inventory
from models.stock_transaction import StockTransaction
from services.stock_ledger import record_transaction

bp = Blueprint('ledger', __name__, url_prefix='/ledger')

_MANUAL_TYPES = {'manual_in', 'manual_out', 'adjust'}
_REASON_REQUIRED = {'manual_out', 'adjust'}


@bp.route('/')
def index():
    return render_template('ledger.html', active_page='ledger')


@bp.route('/api/balances')
def api_balances():
    q = (request.args.get('q') or '').strip()
    query = Inventory.query.filter_by(inventory_type='即納')
    if q:
        like = f'%{q}%'
        query = query.filter(db.or_(
            Inventory.product_code.ilike(like),
            Inventory.product_name.ilike(like),
        ))
    rows = query.order_by(Inventory.product_code).limit(500).all()
    return jsonify({'items': [{
        'product_code': r.product_code,
        'product_name': r.product_name or '',
        'ledger_qty': r.quantity,
        'reserved_qty': r.reserved_qty or 0,
        'available_qty': r.available_qty or 0,
        'location': r.location or '',
    } for r in rows]})


@bp.route('/api/tx', methods=['POST'])
def api_tx():
    body = request.get_json(silent=True) or {}
    product_code = (body.get('product_code') or '').strip()
    tx_type = body.get('tx_type')
    reason = (body.get('reason') or '').strip()
    try:
        qty = int(body.get('qty', 0))
    except (TypeError, ValueError):
        qty = 0

    if not product_code:
        return jsonify({'status': 'error', 'message': '商品コードが必要です'}), 400
    if tx_type not in _MANUAL_TYPES:
        return jsonify({'status': 'error', 'message': f'不正な種別: {tx_type}'}), 400
    if qty <= 0 and tx_type != 'adjust':
        return jsonify({'status': 'error', 'message': '数量は1以上を指定してください'}), 400
    if qty == 0:
        return jsonify({'status': 'error', 'message': '数量に0は指定できません'}), 400
    if tx_type in _REASON_REQUIRED and not reason:
        return jsonify({'status': 'error', 'message': '出庫・調整には理由が必要です'}), 400

    signed_qty = -qty if tx_type == 'manual_out' else qty
    try:
        tx, _ = record_transaction(
            product_code, tx_type, signed_qty, f'manual:{uuid.uuid4()}',
            ref_type='manual', reason=reason or None,
        )
        db.session.commit()
    except ValueError as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 400
    return jsonify({'status': 'ok', 'tx': tx.to_dict()})


@bp.route('/api/history/<path:product_code>')
def api_history(product_code):
    rows = StockTransaction.query.filter_by(product_code=product_code) \
        .order_by(StockTransaction.id.desc()).limit(200).all()
    return jsonify({'items': [r.to_dict() for r in rows]})
