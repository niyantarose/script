"""
自動引当エンジン (Phase 2)
==========================
Yahoo 在庫 (inventory_type='yahoo') を使って受注明細に対し自動的に引当を行う。

呼び出しタイミング:
  1. Yahoo 受注取込後 (import_yahoo_orders) → 新規明細のみ
  2. Yahoo 差分同期後 (yahoo_stock_diff)    → pending / shortage 明細を再試行
  3. Yahoo フル同期後 (yahoo_stock)         → 同上

戻り値:
  {'fully_allocated': n, 'partial': n, 'tori': n, 'shortage': n, 'skipped': n}
"""

from datetime import datetime
from models import db
from models.inventory import Inventory
from models.order_item import OrderItem
from models.allocation import Allocation


# 再試行対象ステータス
RETRY_STATUSES = ('pending', 'shortage', 'partial_waiting')


def run_auto_allocation(item_ids=None):
    """
    自動引当を実行する。

    Parameters
    ----------
    item_ids : list[int] | None
        対象明細 ID リスト。None の場合は RETRY_STATUSES の全明細を処理。

    Returns
    -------
    dict  集計結果
    """
    # ── 対象明細ロード ─────────────────────────────────────────────
    if item_ids is not None:
        items = OrderItem.query.filter(
            OrderItem.id.in_(item_ids),
            OrderItem.status.in_(RETRY_STATUSES),
        ).all()
    else:
        items = OrderItem.query.filter(
            OrderItem.status.in_(RETRY_STATUSES)
        ).all()

    if not items:
        return {'fully_allocated': 0, 'partial': 0, 'tori': 0,
                'shortage': 0, 'skipped': 0}

    # ── 在庫を一括ロード（N+1 回避） ──────────────────────────────
    codes   = list({i.product_code for i in items if i.product_code})
    inv_map: dict[str, Inventory] = {}
    for inv in Inventory.query.filter(
        Inventory.product_code.in_(codes),
        Inventory.inventory_type == 'yahoo',
    ).all():
        inv_map[inv.product_code] = inv

    now   = datetime.now()
    stats = {'fully_allocated': 0, 'partial': 0, 'tori': 0,
             'shortage': 0, 'skipped': 0}

    for item in items:
        needed = item.quantity - (item.allocated_qty or 0)
        if needed <= 0:
            item.status = 'fully_allocated'
            stats['skipped'] += 1
            continue

        inv = inv_map.get(item.product_code)

        # ── ケース①: 即納在庫あり ────────────────────────────────
        if inv and inv.is_immediate:
            avail = max(0, (inv.yahoo_stock or 0) - (inv.reserved_qty or 0))

            if avail >= needed:
                # 全量引当
                inv.reserved_qty  = (inv.reserved_qty or 0) + needed
                inv.available_qty = max(0, (inv.yahoo_stock or 0) - inv.reserved_qty)
                item.allocated_qty = item.quantity
                item.status        = 'fully_allocated'
                _add_alloc(item.id, needed, '本引当', now)
                stats['fully_allocated'] += 1

            elif avail > 0:
                # 部分引当（残りは後続同期で再試行される）
                inv.reserved_qty  = (inv.reserved_qty or 0) + avail
                inv.available_qty = 0
                item.allocated_qty = (item.allocated_qty or 0) + avail
                item.status        = 'partial_waiting'
                _add_alloc(item.id, avail, '本引当', now)
                stats['partial'] += 1

            else:
                # 即納フラグあるが予約で空きゼロ
                item.status = 'shortage'
                stats['shortage'] += 1

        # ── ケース②: お取り寄せ（yahoo_stock>0 だが allow_overdraft=1） ──
        elif inv and (inv.yahoo_stock or 0) > 0:
            item.status = 'provisional_allocated'
            stats['tori'] += 1

        # ── ケース③: 在庫なし ───────────────────────────────────
        else:
            item.status = 'shortage'
            stats['shortage'] += 1

    # commit は呼び出し側で行う
    return stats


def _add_alloc(order_item_id: int, qty: int, alloc_type: str, now: datetime):
    """Allocation レコードを追加（DB セッションに add のみ、commit しない）"""
    db.session.add(Allocation(
        order_item_id   = order_item_id,
        inventory_type  = 'yahoo',
        allocation_type = alloc_type,
        quantity        = qty,
        allocated_by    = 'auto',
        allocated_at    = now,
    ))
