import csv
import io
from flask import Blueprint, render_template, request, jsonify, Response
from models import db
from models.ems import Ems
from models.ems_item import EmsItem
from models.order import Order
from models.order_item import OrderItem
from datetime import datetime, date as date_type

bp = Blueprint('ems', __name__, url_prefix='/ems')

AGENT_LABELS = {'daniel': 'ダニエル', 'tegu': 'テグ'}
PER_PAGE = 20   # 1ページあたりのEMS箱数


def _auto_arrive():
    """到着予定日を過ぎた輸送中EMSを自動的に入荷済みにする。"""
    today = date_type.today()
    targets = Ems.query.filter(
        Ems.status == 'in_transit',
        Ems.estimated_arrival != None,
        Ems.estimated_arrival < today,
    ).all()
    for e in targets:
        e.status     = 'arrived'
        e.arrived_at = e.estimated_arrival
        e.updated_at = datetime.now()
    if targets:
        db.session.commit()
    return len(targets)


def _build_ems_boxes(agent, q='', status_filter='', page=1):
    """EMS箱単位でグループ化したデータを返す。
    戻り値: (boxes_list, total_ems_count)
    """
    _auto_arrive()
    today = date_type.today()

    # ── EMS一覧クエリ（ページネーション） ──────────────────────
    ems_q = Ems.query.filter(Ems.agent == agent)
    if status_filter:
        ems_q = ems_q.filter(Ems.status == status_filter)

    # 検索: EMS番号 or 商品コードが含まれるEMSを絞り込む
    if q:
        matched_ems_ids = db.session.query(EmsItem.ems_id).filter(
            db.or_(
                EmsItem.product_code.contains(q),
                EmsItem.product_name.contains(q),
            )
        ).distinct().subquery()
        ems_q = ems_q.filter(
            db.or_(
                Ems.ems_number.contains(q),
                Ems.id.in_(matched_ems_ids),
            )
        )

    total = ems_q.count()
    ems_list = (
        ems_q.order_by(Ems.shipped_at.desc(), Ems.id.desc())
        .offset((page - 1) * PER_PAGE)
        .limit(PER_PAGE)
        .all()
    )
    if not ems_list:
        return [], total

    ems_ids = [e.id for e in ems_list]

    # ── 対象EMSのアイテムを一括取得（N+1なし） ────────────────
    item_rows = (
        db.session.query(EmsItem, OrderItem, Order)
        .outerjoin(OrderItem, EmsItem.order_item_id == OrderItem.id)
        .outerjoin(Order,     OrderItem.order_id    == Order.id)
        .filter(EmsItem.ems_id.in_(ems_ids))
        .order_by(EmsItem.ems_id, EmsItem.id)
        .all()
    )

    # Python でEMS単位にグループ化
    items_by_ems = {e.id: [] for e in ems_list}
    for ei, oi, o in item_rows:
        items_by_ems[ei.ems_id].append({
            'ems_item':   ei,
            'order_item': oi,
            'order':      o,
        })

    boxes = []
    for e in ems_list:
        days = (today - e.shipped_at).days if e.shipped_at else 0
        is_overdue = (e.status == 'in_transit'
                      and e.arrived_at is None
                      and days >= 10)
        boxes.append({
            'ems':             e,
            'rows':            items_by_ems[e.id],
            'is_overdue':      is_overdue,
            'days_in_transit': days,
        })

    return boxes, total


@bp.route('/')
def index():
    agent         = request.args.get('agent', 'daniel')
    q             = request.args.get('q', '').strip()
    status_filter = request.args.get('status', '')
    page          = max(1, int(request.args.get('page', 1)))

    boxes, total = _build_ems_boxes(agent, q, status_filter, page)
    total_pages  = max(1, (total + PER_PAGE - 1) // PER_PAGE)

    return render_template(
        'ems.html',
        boxes=boxes,
        agent=agent,
        agent_label=AGENT_LABELS.get(agent, agent),
        q=q, status_filter=status_filter,
        page=page, total_pages=total_pages, total=total,
    )


@bp.route('/update/ems/<int:ems_id>', methods=['POST'])
def update_ems(ems_id):
    e = Ems.query.get_or_404(ems_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    allowed = {'shipped_at', 'estimated_arrival', 'arrived_at', 'status', 'memo'}
    if field not in allowed:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        was_arrived = (e.status == 'arrived')

        if field in ('shipped_at', 'estimated_arrival', 'arrived_at'):
            if value:
                setattr(e, field, datetime.strptime(value, '%Y-%m-%d').date())
            else:
                setattr(e, field, None)
            if field == 'arrived_at' and value:
                e.status = 'arrived'
            elif field == 'arrived_at' and not value:
                e.status = 'in_transit'
        else:
            setattr(e, field, value)
        e.updated_at = datetime.now()
        db.session.commit()

        # ── EMS が新たに入荷済みになったら引当 ──────────────────────
        just_arrived = (not was_arrived and e.status == 'arrived')
        if just_arrived or (field == 'status' and value == 'arrived'):
            try:
                from services.allocation import run_ems_arrived_allocation, run_auto_allocation
                # ① EMS内の purchase_no 経由で精確に引当（メイン）
                ems_alloc = run_ems_arrived_allocation(ems_id)
                # ② pending/shortage 全件も再試行（汎用）
                alloc = run_auto_allocation()
                db.session.commit()
                return jsonify({'ok': True,
                                'ems_alloc': ems_alloc,
                                'alloc': alloc})
            except Exception as ae:
                pass  # 引当失敗でも保存は成功

        return jsonify({'ok': True})
    except Exception as ex:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(ex)}), 500


@bp.route('/update/item/<int:item_id>', methods=['POST'])
def update_item(item_id):
    ei = EmsItem.query.get_or_404(item_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    if field not in {'quantity', 'product_code', 'product_name'}:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        if field == 'quantity':
            ei.quantity = int(value)
        else:
            setattr(ei, field, value)
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as ex:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(ex)}), 500


@bp.route('/csv')
def export_csv():
    agent  = request.args.get('agent', 'daniel')
    boxes, _ = _build_ems_boxes(agent, page=1)   # CSV は全件: ページ数で回す
    # 全件取得（ページネーションなし）
    all_boxes = []
    page = 1
    while True:
        b, total = _build_ems_boxes(agent, page=page)
        if not b:
            break
        all_boxes.extend(b)
        if page * PER_PAGE >= total:
            break
        page += 1

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['EMS番号', '発送日', '到着予定日', '入荷日', 'ステータス',
                     '商品コード', '商品名', '数量', '引当先受注番号'])
    for box in all_boxes:
        e = box['ems']
        for row in box['rows']:
            ei = row['ems_item']
            o  = row['order']
            writer.writerow([
                e.ems_number,
                e.shipped_at.strftime('%Y/%m/%d')          if e.shipped_at          else '',
                e.estimated_arrival.strftime('%Y/%m/%d')   if e.estimated_arrival   else '',
                e.arrived_at.strftime('%Y/%m/%d')          if e.arrived_at          else '',
                e.status_label,
                ei.product_code if ei else '',
                ei.product_name if ei else '',
                ei.quantity     if ei else '',
                o.yahoo_order_id if o else (e.order_id or ''),
            ])
    output.seek(0)
    label = AGENT_LABELS.get(agent, agent)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition':
                 f'attachment; filename={label}_ems_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
