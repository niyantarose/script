from __future__ import annotations

import csv
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import StringIO

from flask import current_app

from .constants import ALERT_TYPE_LABELS
from .extensions import db
from .models import (
    Alert,
    Allocation,
    Ems,
    EmsItem,
    Inventory,
    JapanInventoryStaging,
    Order,
    OrderItem,
    Purchase,
)


FINAL_ALLOCATION_TYPES = {"本引当", "手動"}
AUTO_ALERT_TYPES = {
    "purchase_missing",
    "korea_ship_missing",
    "japan_arrival_missing",
    "japan_ship_missing",
    "stock_shortage",
    "delay_warning",
}


def normalize_text(value: str | None) -> str | None:
    if value is None:
        return None
    text = value.strip()
    return text or None


def product_query(product_code: str, product_sub_code: str | None, inventory_type: str):
    query = Inventory.query.filter_by(product_code=product_code, inventory_type=inventory_type)
    if product_sub_code is None:
        query = query.filter(Inventory.product_sub_code.is_(None))
    else:
        query = query.filter_by(product_sub_code=product_sub_code)
    return query


def get_or_create_inventory(product_code: str, product_sub_code: str | None, inventory_type: str) -> Inventory:
    record = product_query(product_code, product_sub_code, inventory_type).first()
    if record:
        return record
    record = Inventory(
        product_code=product_code,
        product_sub_code=product_sub_code,
        inventory_type=inventory_type,
        quantity=0,
        reserved_qty=0,
        available_qty=0,
    )
    db.session.add(record)
    db.session.flush()
    return record


def refresh_inventory_balances() -> None:
    for inventory in Inventory.query.all():
        inventory.available_qty = max(inventory.quantity - inventory.reserved_qty, 0)


def sum_purchase_qty(order_item: OrderItem) -> int:
    return sum(purchase.quantity for purchase in order_item.purchases)


def sum_confirmed_qty(order_item: OrderItem) -> int:
    return sum(
        allocation.quantity
        for allocation in order_item.allocations
        if allocation.allocation_type in FINAL_ALLOCATION_TYPES
    )


def sync_order_statuses() -> None:
    for order in Order.query.all():
        item_statuses = [item.status for item in order.items]
        if item_statuses and all(status == "shipped" for status in item_statuses):
            order.status = "shipped"
        elif item_statuses and all(
            status in {"allocated_sokunou", "fully_allocated", "shipped"}
            for status in item_statuses
        ):
            order.status = "allocated"
        else:
            order.status = "pending"


def recalculate_allocations(commit: bool = True) -> None:
    for inventory in Inventory.query.filter_by(inventory_type="即納").all():
        inventory.reserved_qty = 0

    items = (
        OrderItem.query.join(Order)
        .order_by(Order.ordered_at.asc(), OrderItem.id.asc())
        .all()
    )

    for item in items:
        order = item.order
        confirmed_qty = sum_confirmed_qty(item)
        provisional_qty = sum_purchase_qty(item)
        reserve_qty = 0

        if item.shipped_flag:
            item.allocated_qty = item.quantity
            item.status = "shipped"
            continue

        if order.priority_ship_flag:
            item.allocated_qty = confirmed_qty
            item.status = "priority_hold"
            continue

        remaining_qty = max(item.quantity - confirmed_qty, 0)
        if remaining_qty > 0:
            immediate_inventory = get_or_create_inventory(
                item.product_code,
                item.product_sub_code,
                "即納",
            )
            immediate_available = max(immediate_inventory.quantity - immediate_inventory.reserved_qty, 0)
            reserve_qty = min(remaining_qty, immediate_available)
            immediate_inventory.reserved_qty += reserve_qty

        item.allocated_qty = confirmed_qty + reserve_qty

        if item.allocated_qty >= item.quantity and confirmed_qty == 0:
            item.status = "allocated_sokunou"
        elif confirmed_qty >= item.quantity:
            item.status = "fully_allocated"
        elif reserve_qty > 0:
            item.status = "partial_waiting"
        elif provisional_qty > 0:
            item.status = "provisional_allocated"
        else:
            item.status = "shortage"

    refresh_inventory_balances()
    sync_order_statuses()

    if commit:
        db.session.commit()


def create_purchase(
    order_item_id: int,
    quantity: int,
    shop_name: str | None,
    ordered_at: date,
    status: str,
    memo: str | None,
) -> Purchase:
    item = OrderItem.query.get_or_404(order_item_id)
    purchase = Purchase(
        order_item=item,
        product_code=item.product_code,
        product_sub_code=item.product_sub_code,
        quantity=quantity,
        shop_name=normalize_text(shop_name),
        ordered_at=ordered_at,
        status=status,
        memo=normalize_text(memo),
    )
    db.session.add(purchase)
    db.session.flush()
    db.session.add(
        Allocation(
            order_item=item,
            inventory_type="お取り寄せ",
            allocation_type="仮引当",
            quantity=quantity,
        )
    )
    recalculate_allocations(commit=False)
    db.session.commit()
    return purchase


def create_ems(ems_number: str, shipped_at: date, memo: str | None, items_text: str) -> Ems:
    estimated_arrival = shipped_at + timedelta(days=current_app.config["EMS_LEAD_DAYS"])
    ems = Ems(
        ems_number=ems_number.strip(),
        shipped_at=shipped_at,
        estimated_arrival=estimated_arrival,
        memo=normalize_text(memo),
    )
    db.session.add(ems)
    db.session.flush()

    for line_number, line in enumerate(items_text.splitlines(), start=1):
        if not line.strip():
            continue
        parts = [part.strip() for part in line.split(",")]
        if len(parts) < 2:
            raise ValueError(f"{line_number}行目の形式が不正です。`受注明細ID,数量`で入力してください。")
        order_item_id = int(parts[0])
        quantity = int(parts[1])
        order_item = OrderItem.query.get_or_404(order_item_id)
        ems_item = EmsItem(
            ems=ems,
            order_item=order_item,
            product_code=order_item.product_code,
            product_sub_code=order_item.product_sub_code,
            quantity=quantity,
        )
        db.session.add(ems_item)
        db.session.flush()
        db.session.add(
            Allocation(
                order_item=order_item,
                inventory_type="お取り寄せ",
                allocation_type="本引当",
                quantity=quantity,
                ems_item=ems_item,
            )
        )

    recalculate_allocations(commit=False)
    db.session.commit()
    return ems


def mark_ems_arrived(ems_id: int, arrived_at: date) -> Ems:
    ems = Ems.query.get_or_404(ems_id)
    ems.arrived_at = arrived_at
    ems.status = "arrived"
    for item in ems.items:
        exists = JapanInventoryStaging.query.filter_by(ems_item_id=item.id).first()
        if exists:
            continue
        db.session.add(
            JapanInventoryStaging(
                ems_item=item,
                product_code=item.product_code,
                product_sub_code=item.product_sub_code,
                quantity=item.quantity,
                status="waiting",
            )
        )
    db.session.commit()
    return ems


def manual_allocate(
    order_item_id: int,
    inventory_type: str,
    quantity: int,
    allocated_by: str | None,
) -> Allocation:
    item = OrderItem.query.get_or_404(order_item_id)
    allocated_by = normalize_text(allocated_by)

    if inventory_type == "即納":
        inventory = get_or_create_inventory(item.product_code, item.product_sub_code, "即納")
        refresh_inventory_balances()
        available = max(inventory.quantity - inventory.reserved_qty, 0)
        if available < quantity:
            raise ValueError("即納在庫が不足しています。")
        inventory.quantity -= quantity

    allocation = Allocation(
        order_item=item,
        inventory_type=inventory_type,
        allocation_type="手動",
        quantity=quantity,
        allocated_by=allocated_by,
    )
    db.session.add(allocation)
    recalculate_allocations(commit=False)
    db.session.commit()
    return allocation


def ship_order_item(order_item_id: int) -> OrderItem:
    item = OrderItem.query.get_or_404(order_item_id)
    confirmed_qty = sum_confirmed_qty(item)
    reserved_immediate_qty = max(item.allocated_qty - confirmed_qty, 0)
    if reserved_immediate_qty > 0:
        inventory = get_or_create_inventory(item.product_code, item.product_sub_code, "即納")
        inventory.quantity = max(inventory.quantity - reserved_immediate_qty, 0)
        inventory.reserved_qty = max(inventory.reserved_qty - reserved_immediate_qty, 0)
    item.shipped_flag = True
    item.status = "shipped"
    item.allocated_qty = item.quantity
    recalculate_allocations(commit=False)
    db.session.commit()
    return item


def update_order_meta(
    order_id: int,
    priority_ship_flag: bool,
    delay_memo: str | None,
    customer_contacted_flag: bool,
) -> Order:
    order = Order.query.get_or_404(order_id)
    order.priority_ship_flag = priority_ship_flag
    order.delay_memo = normalize_text(delay_memo)
    order.customer_contacted_flag = customer_contacted_flag
    recalculate_allocations(commit=False)
    db.session.commit()
    return order


def create_alert(alert_type: str, order: Order | None, order_item: OrderItem | None, message: str) -> None:
    db.session.add(
        Alert(
            alert_type=alert_type,
            order=order,
            order_item=order_item,
            product_code=order_item.product_code if order_item else None,
            message=message,
        )
    )


def run_checks() -> None:
    Alert.query.filter(
        Alert.alert_type.in_(AUTO_ALERT_TYPES),
        Alert.resolved_flag.is_(False),
    ).delete(synchronize_session=False)
    db.session.flush()

    delay_days = current_app.config["DELAY_WARNING_DAYS"]
    today = date.today()
    ems_item_order_ids = {ems_item.order_item_id for ems_item in EmsItem.query.all()}

    for item in OrderItem.query.join(Order).all():
        if item.shipped_flag:
            continue
        if not item.purchases and item.status not in {"allocated_sokunou", "fully_allocated", "priority_hold"}:
            create_alert(
                "purchase_missing",
                item.order,
                item,
                f"受注 {item.order.yahoo_order_id} の {item.product_code} に発注記録がありません。",
            )
        if item.status == "shortage":
            create_alert(
                "stock_shortage",
                item.order,
                item,
                f"{item.product_code} の引当在庫が不足しています。",
            )
        order_age = (today - item.order.ordered_date).days
        if order_age >= delay_days and item.order.status != "shipped":
            create_alert(
                "delay_warning",
                item.order,
                item,
                f"受注 {item.order.yahoo_order_id} は注文から {order_age} 日経過しています。",
            )
        if item.status in {"allocated_sokunou", "fully_allocated"} and not item.shipped_flag:
            create_alert(
                "japan_ship_missing",
                item.order,
                item,
                f"受注 {item.order.yahoo_order_id} は発送可能ですが未発送です。",
            )

    for purchase in Purchase.query.all():
        if purchase.status == "arrived" and purchase.order_item_id not in ems_item_order_ids:
            create_alert(
                "korea_ship_missing",
                purchase.order_item.order,
                purchase.order_item,
                f"{purchase.product_code} は韓国入荷済みですが EMS 未登録です。",
            )

    for ems in Ems.query.filter_by(status="in_transit").all():
        if ems.estimated_arrival < today:
            for item in ems.items:
                create_alert(
                    "japan_arrival_missing",
                    item.order_item.order,
                    item.order_item,
                    f"EMS {ems.ems_number} の {item.product_code} が到着予定日超過です。",
                )

    db.session.commit()


def resolve_alert(alert_id: int) -> Alert:
    alert = Alert.query.get_or_404(alert_id)
    alert.resolved_flag = True
    alert.resolved_at = datetime.now()
    db.session.commit()
    return alert


def delay_level_for_order(order: Order) -> str:
    today = date.today()
    age = (today - order.ordered_date).days
    delay_days = current_app.config["DELAY_WARNING_DAYS"]

    if age >= delay_days:
        return "danger"

    for item in order.items:
        if any(purchase.status == "arrived" for purchase in item.purchases) and not item.ems_items:
            return "warning"
        if item.status in {"allocated_sokunou", "fully_allocated"} and not item.shipped_flag:
            return "notice"

    return "ok"


def delivery_judgement(order: Order, ems_shipped_at: date | None = None) -> tuple[str, str]:
    if not order.desired_delivery_date:
        return "未指定", "希望日なし"

    if not ems_shipped_at:
        ems = (
            Ems.query.join(Ems.items)
            .join(EmsItem.order_item)
            .filter(OrderItem.order_id == order.id)
            .order_by(Ems.shipped_at.desc())
            .first()
        )
        ems_shipped_at = ems.shipped_at if ems else None

    if not ems_shipped_at:
        return "保留", "EMS発送日未設定"

    threshold = ems_shipped_at + timedelta(
        days=current_app.config["EMS_LEAD_DAYS"] + 1 + current_app.config["SHIPPING_BUFFER_DAYS"]
    )
    if order.desired_delivery_date >= threshold:
        return "OK", f"{threshold.isoformat()} までに確保可能"
    return "保留", f"{threshold.isoformat()} より希望日が早いため保留"


def dashboard_stats() -> dict:
    alerts = Alert.query.filter_by(resolved_flag=False).all()
    return {
        "ready_orders": Order.query.filter_by(status="allocated").count(),
        "delay_orders": len(
            [order for order in Order.query.all() if delay_level_for_order(order) == "danger"]
        ),
        "purchase_missing": len([alert for alert in alerts if alert.alert_type == "purchase_missing"]),
        "stock_shortage": len([alert for alert in alerts if alert.alert_type == "stock_shortage"]),
        "unresolved_alerts": len(alerts),
        "ems_in_transit": Ems.query.filter_by(status="in_transit").count(),
    }


def latest_orders(limit: int = 8) -> list[Order]:
    return Order.query.order_by(Order.ordered_at.desc()).limit(limit).all()


def export_rows(kind: str) -> tuple[str, list[dict]]:
    if kind == "orders":
        rows = [
            {
                "受注番号": order.yahoo_order_id,
                "注文日時": order.ordered_at.isoformat(sep=" ", timespec="minutes"),
                "顧客名": order.customer_name or "",
                "受注ステータス": order.status,
                "先送り": int(order.priority_ship_flag),
                "連絡済み": int(order.customer_contacted_flag),
            }
            for order in Order.query.order_by(Order.ordered_at.desc()).all()
        ]
    elif kind == "ems":
        rows = [
            {
                "EMS番号": ems.ems_number,
                "発送日": ems.shipped_at.isoformat(),
                "到着予定": ems.estimated_arrival.isoformat(),
                "入荷日": ems.arrived_at.isoformat() if ems.arrived_at else "",
                "ステータス": ems.status,
            }
            for ems in Ems.query.order_by(Ems.shipped_at.desc()).all()
        ]
    elif kind == "checks":
        rows = [
            {
                "アラート種別": ALERT_TYPE_LABELS.get(alert.alert_type, alert.alert_type),
                "受注番号": alert.order.yahoo_order_id if alert.order else "",
                "商品コード": alert.product_code or "",
                "メッセージ": alert.message,
            }
            for alert in Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).all()
        ]
    elif kind == "allocations":
        rows = [
            {
                "受注明細ID": allocation.order_item_id,
                "受注番号": allocation.order_item.order.yahoo_order_id,
                "商品コード": allocation.order_item.product_code,
                "引当種別": allocation.allocation_type,
                "在庫種別": allocation.inventory_type,
                "数量": allocation.quantity,
                "実行者": allocation.allocated_by or "",
                "日時": allocation.allocated_at.isoformat(sep=" ", timespec="minutes"),
            }
            for allocation in Allocation.query.order_by(Allocation.allocated_at.desc()).all()
        ]
    elif kind == "japan_stock":
        rows = [
            {
                "ID": row.id,
                "EMS番号": row.ems_item.ems.ems_number,
                "商品コード": row.product_code,
                "数量": row.quantity,
                "ステータス": row.status,
                "受注明細ID": row.assigned_order_item_id or "",
                "反映日時": row.reflected_at.isoformat(sep=" ", timespec="minutes")
                if row.reflected_at
                else "",
            }
            for row in JapanInventoryStaging.query.order_by(JapanInventoryStaging.created_at.desc()).all()
        ]
    elif kind == "alerts":
        rows = [
            {
                "アラート種別": ALERT_TYPE_LABELS.get(alert.alert_type, alert.alert_type),
                "受注番号": alert.order.yahoo_order_id if alert.order else "",
                "商品コード": alert.product_code or "",
                "解決": int(alert.resolved_flag),
                "内容": alert.message,
            }
            for alert in Alert.query.order_by(Alert.created_at.desc()).all()
        ]
    elif kind == "all_tables":
        counts = {
            "orders": Order.query.count(),
            "order_items": OrderItem.query.count(),
            "purchases": Purchase.query.count(),
            "ems": Ems.query.count(),
            "ems_items": EmsItem.query.count(),
            "inventory": Inventory.query.count(),
            "allocations": Allocation.query.count(),
            "alerts": Alert.query.count(),
            "japan_inventory_staging": JapanInventoryStaging.query.count(),
        }
        rows = [{"テーブル": name, "件数": count} for name, count in counts.items()]
    else:
        raise ValueError("対応していないCSV種別です。")
    return kind, rows


def export_csv(kind: str) -> str:
    _, rows = export_rows(kind)
    if not rows:
        return "\ufeff"

    output = StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return "\ufeff" + output.getvalue()


def update_staging_row(
    staging_id: int,
    action: str,
    assigned_order_item_id: int | None = None,
    excluded_reason: str | None = None,
) -> JapanInventoryStaging:
    row = JapanInventoryStaging.query.get_or_404(staging_id)
    if action == "assign":
        if not assigned_order_item_id:
            raise ValueError("受注明細IDを指定してください。")
        manual_allocate(assigned_order_item_id, "お取り寄せ", row.quantity, "日本在庫仕分け")
        row.assigned_order_item_id = assigned_order_item_id
        row.status = "assigned_to_order"
    elif action == "stock":
        row.status = "to_japan_stock"
    elif action == "exclude":
        row.status = "excluded"
        row.excluded_reason = normalize_text(excluded_reason)
    elif action == "return":
        row.status = "returned_to_ems"
    else:
        raise ValueError("不正な仕分けアクションです。")
    db.session.commit()
    return row


def reflect_japan_stock() -> int:
    rows = JapanInventoryStaging.query.filter_by(status="to_japan_stock").all()
    if not rows:
        return 0

    grouped: dict[tuple[str, str | None], int] = defaultdict(int)
    for row in rows:
        grouped[(row.product_code, row.product_sub_code)] += row.quantity

    for (product_code, product_sub_code), quantity in grouped.items():
        inventory = get_or_create_inventory(product_code, product_sub_code, "即納")
        inventory.quantity += quantity

    now = datetime.now()
    for row in rows:
        row.status = "reflected"
        row.reflected_at = now

    recalculate_allocations(commit=False)
    db.session.commit()
    return len(rows)
