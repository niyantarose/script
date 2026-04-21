"""
在庫リアルタイム確認画面
========================
Yahoo から同期した在庫を検索・閲覧するためのルート。
30 秒ごとに自動更新し、即納 / お取り寄せ / 在庫なし の状態を視覚的に表示。
"""

from flask import Blueprint, render_template, request, jsonify
from models.inventory import Inventory
from models import db
from sqlalchemy import func, case, or_

bp = Blueprint('stock', __name__)


# ── ページ ────────────────────────────────────────────────────────────

@bp.route('/stock')
def index():
    return render_template('stock.html', active_page='stock')


# ── JSON API ─────────────────────────────────────────────────────────

@bp.route('/stock/api')
def api():
    """
    GET /stock/api
      ?q=<検索ワード(商品コード/商品名)>
      &filter=all|sokunou|tori|zero
      &page=1
      &per=200
    """
    q           = request.args.get('q', '').strip()
    filter_type = request.args.get('filter', 'all')   # all / sokunou / tori / zero
    page        = max(1, int(request.args.get('page', 1)))
    per_page    = min(500, max(10, int(request.args.get('per', 200))))

    # ── 一覧クエリ ──────────────────────────────────────────────────
    base_q = Inventory.query.filter(Inventory.inventory_type == 'yahoo')

    if q:
        like = f'%{q}%'
        base_q = base_q.filter(
            or_(
                Inventory.product_code.ilike(like),
                Inventory.product_name.ilike(like),
            )
        )

    if filter_type == 'sokunou':
        base_q = base_q.filter(Inventory.is_immediate == True)
    elif filter_type == 'tori':
        base_q = base_q.filter(
            Inventory.is_immediate == False,
            Inventory.quantity > 0,
        )
    elif filter_type == 'zero':
        base_q = base_q.filter(Inventory.quantity <= 0)

    total = base_q.count()
    items = (
        base_q
        .order_by(Inventory.is_immediate.desc(), Inventory.quantity.desc(), Inventory.product_code)
        .offset((page - 1) * per_page)
        .limit(per_page)
        .all()
    )

    # ── サマリー統計（1クエリで全カウント） ─────────────────────────
    stats_row = db.session.query(
        func.count().label('total'),
        func.sum(case((Inventory.is_immediate == True,  1), else_=0)).label('sokunou'),
        func.sum(case((
            (Inventory.is_immediate == False) & (Inventory.quantity > 0),
            1), else_=0)).label('tori'),
        func.sum(case((Inventory.quantity <= 0, 1), else_=0)).label('zero'),
        func.max(Inventory.last_synced_at).label('last_sync'),
    ).filter(Inventory.inventory_type == 'yahoo').first()

    last_synced = (
        stats_row.last_sync.strftime('%Y/%m/%d %H:%M:%S')
        if stats_row and stats_row.last_sync else 'なし'
    )

    return jsonify({
        'items':    [_item_dict(i) for i in items],
        'total':    total,
        'page':     page,
        'per_page': per_page,
        'pages':    max(1, (total + per_page - 1) // per_page),
        'stats': {
            'total':       stats_row.total   if stats_row else 0,
            'sokunou':     stats_row.sokunou if stats_row else 0,
            'tori':        stats_row.tori    if stats_row else 0,
            'zero':        stats_row.zero    if stats_row else 0,
            'last_synced': last_synced,
        },
    })


def _item_dict(inv):
    yahoo  = inv.yahoo_stock  or 0
    resv   = inv.reserved_qty or 0
    avail  = inv.available_qty if inv.available_qty is not None else max(0, yahoo - resv)
    return {
        'id':            inv.id,
        'product_code':  inv.product_code,
        'sub_code':      inv.product_sub_code or '',
        'product_name':  inv.product_name  or '',
        'yahoo_stock':   yahoo,
        'reserved_qty':  resv,
        'available_qty': avail,
        'is_immediate':  inv.is_immediate,
        'last_synced_at': (
            inv.last_synced_at.strftime('%m/%d %H:%M')
            if inv.last_synced_at else ''
        ),
    }
