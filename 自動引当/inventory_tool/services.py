from __future__ import annotations

import csv
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import StringIO

from flask import current_app
from sqlalchemy import or_

from .constants import (
    ALERT_TYPE_LABELS,
    EMS_STATUS_LABELS,
    JAPAN_STAGING_STATUS_LABELS,
    ORDER_ITEM_STATUS_LABELS,
    ORDER_STATUS_LABELS,
    PURCHASE_STATUS_LABELS,
    SOURCE_TYPE_LABELS,
)
from .extensions import db
from .models import (
    Alert,
    Allocation,
    EditLog,
    Ems,
    EmsItem,
    ImportedFile,
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
VALID_SOURCE_TYPES = set(SOURCE_TYPE_LABELS)
VALID_PURCHASE_STATUS = set(PURCHASE_STATUS_LABELS)
VALID_EMS_STATUS = set(EMS_STATUS_LABELS)
VALID_ORDER_ITEM_STATUS = set(ORDER_ITEM_STATUS_LABELS)


def normalize_text(value: str | None) -> str | None:
    if value is None:
        return None
    text = value.strip()
    return text or None


def normalize_source_type(source_type: str | None, default: str = "daniel") -> str:
    value = normalize_text(source_type) or default
    if value not in VALID_SOURCE_TYPES:
        raise ValueError("対応していないデータ種別です。")
    return value


def _stringify_value(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def _parse_date_value(value: str | None, nullable: bool = True) -> date | None:
    text = normalize_text(value)
    if text is None:
        if nullable:
            return None
        raise ValueError("日付を入力してください。")
    return date.fromisoformat(text)


def _parse_int_value(value: str | None, nullable: bool = False) -> int | None:
    text = normalize_text(value)
    if text is None:
        if nullable:
            return None
        raise ValueError("数値を入力してください。")
    return int(text)


def _parse_bool_value(value) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    return text in {"1", "true", "yes", "on"}


def _log_edit(table_name: str, record_id: int, field_name: str, old_value, new_value, edited_by: str | None) -> None:
    db.session.add(
        EditLog(
            table_name=table_name,
            record_id=record_id,
            field_name=field_name,
            old_value=_stringify_value(old_value),
            new_value=_stringify_value(new_value),
            edited_by=normalize_text(edited_by) or "web",
        )
    )


def _ensure_staging_rows_for_ems(ems: Ems) -> None:
    if ems.arrived_at is None:
        return
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
    source_type: str = "daniel",
) -> Purchase:
    item = OrderItem.query.get_or_404(order_item_id)
    purchase = Purchase(
        order_item=item,
        source_type=normalize_source_type(source_type),
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


def create_ems(
    ems_number: str,
    shipped_at: date,
    memo: str | None,
    items_text: str,
    source_type: str = "daniel",
) -> Ems:
    estimated_arrival = shipped_at + timedelta(days=current_app.config["EMS_LEAD_DAYS"])
    ems = Ems(
        source_type=normalize_source_type(source_type),
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
    _ensure_staging_rows_for_ems(ems)
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
        has_active_purchase = any(purchase.status != "pending_order" for purchase in item.purchases)
        if not has_active_purchase and item.status not in {"allocated_sokunou", "fully_allocated", "priority_hold"}:
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
        if purchase.status in {"arrived", "shipped"} and purchase.order_item_id not in ems_item_order_ids:
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


def latest_ems_for_item(order_item: OrderItem) -> Ems | None:
    ems_list = [
        ems_item.ems
        for ems_item in order_item.ems_items
        if ems_item.ems is not None
    ]
    if not ems_list:
        return None
    return max(ems_list, key=lambda ems: (ems.shipped_at or date.min, ems.id or 0))


def latest_arrived_ems_number(order_item: OrderItem) -> str:
    arrived_ems = [
        ems_item.ems
        for ems_item in order_item.ems_items
        if ems_item.ems is not None and ems_item.ems.arrived_at is not None
    ]
    if not arrived_ems:
        return "-"
    latest = max(arrived_ems, key=lambda ems: (ems.arrived_at or date.min, ems.id or 0))
    return latest.ems_number


def delivery_status_for_item(order_item: OrderItem) -> str:
    if order_item.shipped_flag:
        return "発送済み"
    if order_item.inventory_type == "即納" or order_item.status == "allocated_sokunou":
        return "即納（日本在庫）"

    latest_ems = latest_ems_for_item(order_item)
    if latest_ems is not None:
        if latest_ems.arrived_at:
            return "EMS入荷済み"
        return "お取り寄せ（韓国から）"

    if any(purchase.status == "arrived" for purchase in order_item.purchases):
        return "韓国入荷済み"
    if order_item.purchases:
        return "お取り寄せ（発注済み）"
    return "未手配"


def purchase_status_for_item(order_item: OrderItem) -> str:
    if any(purchase.status == "shipped" for purchase in order_item.purchases):
        return "発送済"
    if any(purchase.status == "arrived" for purchase in order_item.purchases):
        return "入荷済"
    if any(purchase.status == "ordered" for purchase in order_item.purchases):
        return "発注済"
    return "未発注"


def update_record_field(
    entity: str,
    record_id: int,
    field_name: str,
    value,
    edited_by: str | None = None,
):
    entity = normalize_text(entity) or ""
    field_name = normalize_text(field_name) or ""

    config = {
        "order": {
            "model": Order,
            "table": "orders",
            "fields": {
                "yahoo_order_id": {"type": "text"},
                "customer_code": {"type": "text", "nullable": True},
                "customer_name": {"type": "text", "nullable": True},
                "desired_delivery_date": {"type": "date", "nullable": True},
                "delay_memo": {"type": "text", "nullable": True},
                "priority_ship_flag": {"type": "bool"},
                "customer_contacted_flag": {"type": "bool"},
                "status": {"type": "choice", "choices": ORDER_STATUS_LABELS},
            },
            "recalculate": False,
        },
        "order_item": {
            "model": OrderItem,
            "table": "order_items",
            "fields": {
                "product_code": {"type": "text", "recalculate": True},
                "product_sub_code": {"type": "text", "nullable": True, "recalculate": True},
                "quantity": {"type": "int", "recalculate": True},
                "inventory_type": {
                    "type": "choice",
                    "choices": {"即納": "即納", "お取り寄せ": "お取り寄せ"},
                    "recalculate": True,
                },
                "status": {"type": "choice", "choices": ORDER_ITEM_STATUS_LABELS},
            },
            "recalculate": False,
        },
        "purchase": {
            "model": Purchase,
            "table": "purchases",
            "fields": {
                "product_code": {"type": "text"},
                "product_sub_code": {"type": "text", "nullable": True},
                "quantity": {"type": "int"},
                "shop_name": {"type": "text", "nullable": True},
                "ordered_at": {"type": "date"},
                "status": {"type": "choice", "choices": PURCHASE_STATUS_LABELS},
                "memo": {"type": "text", "nullable": True},
            },
            "recalculate": False,
        },
        "ems": {
            "model": Ems,
            "table": "ems",
            "fields": {
                "ems_number": {"type": "text"},
                "shipped_at": {"type": "date"},
                "arrived_at": {"type": "date", "nullable": True},
                "status": {"type": "choice", "choices": EMS_STATUS_LABELS},
                "memo": {"type": "text", "nullable": True},
            },
            "recalculate": False,
        },
        "japan_stock": {
            "model": JapanInventoryStaging,
            "table": "japan_inventory_staging",
            "fields": {
                "product_code": {"type": "text"},
                "product_sub_code": {"type": "text", "nullable": True},
                "quantity": {"type": "int"},
                "status": {"type": "choice", "choices": JAPAN_STAGING_STATUS_LABELS},
                "assigned_order_item_id": {"type": "int", "nullable": True},
                "excluded_reason": {"type": "text", "nullable": True},
            },
            "recalculate": False,
        },
    }

    if entity not in config:
        raise ValueError("編集対象が不正です。")

    entity_config = config[entity]
    if field_name not in entity_config["fields"]:
        raise ValueError("この項目は編集できません。")

    record = entity_config["model"].query.get_or_404(record_id)
    field_config = entity_config["fields"][field_name]
    old_value = getattr(record, field_name)

    field_type = field_config["type"]
    nullable = field_config.get("nullable", False)
    if field_type == "text":
        parsed_value = normalize_text(value) if nullable else (normalize_text(value) or "")
        if not nullable and not parsed_value:
            raise ValueError("値を入力してください。")
    elif field_type == "date":
        parsed_value = _parse_date_value(value, nullable=nullable)
    elif field_type == "int":
        parsed_value = _parse_int_value(value, nullable=nullable)
    elif field_type == "bool":
        parsed_value = _parse_bool_value(value)
    elif field_type == "choice":
        parsed_value = normalize_text(value) or ""
        if parsed_value not in field_config["choices"]:
            raise ValueError("選択値が不正です。")
    else:
        raise ValueError("未対応の編集種別です。")

    if _stringify_value(old_value) == _stringify_value(parsed_value):
        return record

    setattr(record, field_name, parsed_value)

    if entity == "ems":
        if field_name == "shipped_at" and record.shipped_at:
            record.estimated_arrival = record.shipped_at + timedelta(days=current_app.config["EMS_LEAD_DAYS"])
        if field_name == "status" and record.status == "arrived" and record.arrived_at is None:
            record.arrived_at = date.today()
        if field_name == "arrived_at" and record.arrived_at and record.status == "in_transit":
            record.status = "arrived"
        _ensure_staging_rows_for_ems(record)

    _log_edit(entity_config["table"], record.id, field_name, old_value, parsed_value, edited_by)

    should_recalculate = field_config.get("recalculate", entity_config["recalculate"])
    if should_recalculate:
        recalculate_allocations(commit=False)
    db.session.commit()
    run_checks()
    return record


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
        return "引当可", f"{threshold.isoformat()} までに確保可能"
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


def latest_purchases(source_type: str, limit: int = 5) -> list[Purchase]:
    return (
        Purchase.query.filter_by(source_type=normalize_source_type(source_type))
        .order_by(Purchase.ordered_at.desc(), Purchase.id.desc())
        .limit(limit)
        .all()
    )


def latest_ems(source_type: str, limit: int = 5) -> list[Ems]:
    return (
        Ems.query.filter_by(source_type=normalize_source_type(source_type))
        .order_by(Ems.shipped_at.desc(), Ems.id.desc())
        .limit(limit)
        .all()
    )


def order_status_counts() -> list[dict[str, int | str]]:
    return [
        {"key": key, "label": label, "count": Order.query.filter_by(status=key).count()}
        for key, label in ORDER_STATUS_LABELS.items()
    ]


def search_orders_data(
    order_keyword: str | None = None,
    customer_keyword: str | None = None,
    product_keyword: str | None = None,
) -> list[Order]:
    order_keyword = normalize_text(order_keyword)
    customer_keyword = normalize_text(customer_keyword)
    product_keyword = normalize_text(product_keyword)

    query = Order.query
    if product_keyword:
        query = query.join(Order.items)

    if order_keyword:
        query = query.filter(Order.yahoo_order_id.ilike(f"%{order_keyword}%"))
    if customer_keyword:
        query = query.filter(
            or_(
                Order.customer_code.ilike(f"%{customer_keyword}%"),
                Order.customer_name.ilike(f"%{customer_keyword}%"),
            )
        )
    if product_keyword:
        query = query.filter(
            or_(
                OrderItem.product_code.ilike(f"%{product_keyword}%"),
                OrderItem.product_sub_code.ilike(f"%{product_keyword}%"),
            )
        )
        query = query.distinct()

    return query.order_by(Order.ordered_at.desc()).all()


def _ensure_import_inventory() -> int:
    count = 0
    for row in [
        {"product_code": "SKU-1001", "product_sub_code": None, "inventory_type": "即納", "quantity": 5},
        {"product_code": "SKU-1002", "product_sub_code": "BLUE", "inventory_type": "即納", "quantity": 1},
        {"product_code": "SKU-2001", "product_sub_code": None, "inventory_type": "お取り寄せ", "quantity": 99},
    ]:
        inventory = product_query(
            row["product_code"],
            row["product_sub_code"],
            row["inventory_type"],
        ).first()
        if inventory:
            continue
        db.session.add(
            Inventory(
                product_code=row["product_code"],
                product_sub_code=row["product_sub_code"],
                inventory_type=row["inventory_type"],
                quantity=row["quantity"],
                reserved_qty=0,
                available_qty=row["quantity"],
            )
        )
        count += 1
    db.session.flush()
    return count


def _ensure_import_orders() -> dict[str, Order]:
    now = datetime.now()
    created_orders: dict[str, Order] = {}
    samples = [
        {
            "yahoo_order_id": "YH-20260413-001",
            "ordered_at": now - timedelta(days=1),
            "desired_delivery_date": date.today() + timedelta(days=6),
            "customer_code": "CUS-1001",
            "customer_name": "田中 花子",
            "items": [
                {"product_code": "SKU-1001", "product_sub_code": None, "quantity": 2, "inventory_type": "即納"},
            ],
        },
        {
            "yahoo_order_id": "YH-20260413-002",
            "ordered_at": now - timedelta(days=3),
            "desired_delivery_date": date.today() + timedelta(days=5),
            "customer_code": "CUS-1002",
            "customer_name": "佐藤 次郎",
            "items": [
                {
                    "product_code": "SKU-1002",
                    "product_sub_code": "BLUE",
                    "quantity": 2,
                    "inventory_type": "お取り寄せ",
                },
            ],
        },
        {
            "yahoo_order_id": "YH-20260413-003",
            "ordered_at": now - timedelta(days=7),
            "desired_delivery_date": date.today() + timedelta(days=3),
            "customer_code": "CUS-1003",
            "customer_name": "鈴木 一郎",
            "delay_memo": "仕入先確認中",
            "items": [
                {"product_code": "SKU-3001", "product_sub_code": None, "quantity": 1, "inventory_type": "お取り寄せ"},
            ],
        },
        {
            "yahoo_order_id": "YH-20260413-004",
            "ordered_at": now - timedelta(days=4),
            "desired_delivery_date": date.today() + timedelta(days=7),
            "customer_code": "CUS-1004",
            "customer_name": "高橋 美咲",
            "items": [
                {"product_code": "SKU-2001", "product_sub_code": None, "quantity": 1, "inventory_type": "お取り寄せ"},
            ],
        },
        {
            "yahoo_order_id": "YH-20260413-005",
            "ordered_at": now - timedelta(days=2),
            "desired_delivery_date": date.today() + timedelta(days=8),
            "customer_code": "CUS-1005",
            "customer_name": "山田 優",
            "items": [
                {"product_code": "SKU-4001", "product_sub_code": None, "quantity": 2, "inventory_type": "お取り寄せ"},
            ],
        },
        {
            "yahoo_order_id": "YH-20260413-006",
            "ordered_at": now - timedelta(days=5),
            "desired_delivery_date": date.today() + timedelta(days=9),
            "customer_code": "CUS-1006",
            "customer_name": "木村 愛",
            "items": [
                {"product_code": "SKU-5001", "product_sub_code": "SET", "quantity": 1, "inventory_type": "お取り寄せ"},
            ],
        },
    ]

    for sample in samples:
        order = Order.query.filter_by(yahoo_order_id=sample["yahoo_order_id"]).first()
        if order:
            if not order.customer_code and sample.get("customer_code"):
                order.customer_code = sample["customer_code"]
            created_orders[sample["yahoo_order_id"]] = order
            continue

        order = Order(
            yahoo_order_id=sample["yahoo_order_id"],
            ordered_at=sample["ordered_at"],
            desired_delivery_date=sample["desired_delivery_date"],
            customer_code=sample.get("customer_code"),
            customer_name=sample.get("customer_name"),
            status="pending",
            delay_memo=sample.get("delay_memo"),
        )
        order.items = [
            OrderItem(
                product_code=item["product_code"],
                product_sub_code=item["product_sub_code"],
                quantity=item["quantity"],
                inventory_type=item["inventory_type"],
            )
            for item in sample["items"]
        ]
        db.session.add(order)
        db.session.flush()
        created_orders[sample["yahoo_order_id"]] = order

    return created_orders


def _record_imported_file(file_name: str, file_type: str, record_count: int) -> bool:
    exists = ImportedFile.query.filter_by(file_name=file_name).first()
    if exists:
        return False
    db.session.add(ImportedFile(file_name=file_name, file_type=file_type, record_count=record_count))
    return True


def _ensure_import_purchases(source_type: str) -> int:
    source_type = normalize_source_type(source_type)
    _ensure_import_inventory()
    orders = _ensure_import_orders()
    purchase_specs = [
        {
            "source_type": "daniel",
            "order_id": "YH-20260413-002",
            "product_code": "SKU-1002",
            "product_sub_code": "BLUE",
            "quantity": 1,
            "shop_name": "K-Shop",
            "ordered_at": date.today() - timedelta(days=1),
            "status": "arrived",
        },
        {
            "source_type": "daniel",
            "order_id": "YH-20260413-004",
            "product_code": "SKU-2001",
            "product_sub_code": None,
            "quantity": 1,
            "shop_name": "K-Mall",
            "ordered_at": date.today() - timedelta(days=2),
            "status": "arrived",
        },
        {
            "source_type": "tegu",
            "order_id": "YH-20260413-005",
            "product_code": "SKU-4001",
            "product_sub_code": None,
            "quantity": 2,
            "shop_name": "テグ発注リスト",
            "ordered_at": date.today() - timedelta(days=1),
            "status": "ordered",
        },
        {
            "source_type": "tegu",
            "order_id": "YH-20260413-006",
            "product_code": "SKU-5001",
            "product_sub_code": "SET",
            "quantity": 1,
            "shop_name": "テグ発注リスト",
            "ordered_at": date.today() - timedelta(days=3),
            "status": "arrived",
        },
    ]
    created = 0

    for spec in purchase_specs:
        if spec["source_type"] != source_type:
            continue
        order = orders[spec["order_id"]]
        item = next(
            (
                row
                for row in order.items
                if row.product_code == spec["product_code"] and row.product_sub_code == spec["product_sub_code"]
            ),
            None,
        )
        if item is None:
            continue

        exists = Purchase.query.filter_by(
            order_item_id=item.id,
            source_type=spec["source_type"],
            product_code=spec["product_code"],
            product_sub_code=spec["product_sub_code"],
            quantity=spec["quantity"],
            shop_name=spec["shop_name"],
            ordered_at=spec["ordered_at"],
        ).first()
        if exists:
            continue

        db.session.add(
            Purchase(
                order_item=item,
                source_type=spec["source_type"],
                product_code=spec["product_code"],
                product_sub_code=spec["product_sub_code"],
                quantity=spec["quantity"],
                shop_name=spec["shop_name"],
                ordered_at=spec["ordered_at"],
                status=spec["status"],
            )
        )
        db.session.flush()
        db.session.add(
            Allocation(
                order_item=item,
                inventory_type="お取り寄せ",
                allocation_type="仮引当",
                quantity=spec["quantity"],
            )
        )
        created += 1

    return created


def _ensure_import_ems(source_type: str) -> int:
    source_type = normalize_source_type(source_type)
    _ensure_import_purchases(source_type)
    orders = _ensure_import_orders()
    ems_specs = [
        {
            "source_type": "daniel",
            "ems_number": "EMS123456789KR",
            "order_id": "YH-20260413-004",
            "product_code": "SKU-2001",
            "product_sub_code": None,
            "quantity": 1,
            "shipped_at": date.today() - timedelta(days=4),
            "arrived_at": date.today() - timedelta(days=1),
            "status": "arrived",
            "memo": "ダニエル4月便",
        },
        {
            "source_type": "tegu",
            "ems_number": "EMS987654321KR",
            "order_id": "YH-20260413-006",
            "product_code": "SKU-5001",
            "product_sub_code": "SET",
            "quantity": 1,
            "shipped_at": date.today() - timedelta(days=2),
            "arrived_at": None,
            "status": "in_transit",
            "memo": "テグ4月便",
        },
    ]
    created = 0

    for spec in ems_specs:
        if spec["source_type"] != source_type:
            continue

        order = orders.get(spec["order_id"])
        if order is None:
            continue
        item = next(
            (
                row
                for row in order.items
                if row.product_code == spec["product_code"] and row.product_sub_code == spec["product_sub_code"]
            ),
            None,
        )
        if item is None:
            continue

        ems = Ems.query.filter_by(ems_number=spec["ems_number"]).first()
        if ems is None:
            ems = Ems(
                source_type=spec["source_type"],
                ems_number=spec["ems_number"],
                shipped_at=spec["shipped_at"],
                estimated_arrival=spec["shipped_at"] + timedelta(days=current_app.config["EMS_LEAD_DAYS"]),
                arrived_at=spec["arrived_at"],
                status=spec["status"],
                memo=spec["memo"],
            )
            db.session.add(ems)
            db.session.flush()
            created += 1
        elif not ems.source_type:
            ems.source_type = spec["source_type"]

        ems_item = EmsItem.query.filter_by(ems_id=ems.id, order_item_id=item.id).first()
        if ems_item is None:
            ems_item = EmsItem(
                ems=ems,
                order_item=item,
                product_code=item.product_code,
                product_sub_code=item.product_sub_code,
                quantity=spec["quantity"],
            )
            db.session.add(ems_item)
            db.session.flush()
            created += 1

        final_allocation = Allocation.query.filter_by(
            order_item_id=item.id,
            allocation_type="本引当",
            ems_item_id=ems_item.id,
        ).first()
        if final_allocation is None:
            db.session.add(
                Allocation(
                    order_item=item,
                    inventory_type="お取り寄せ",
                    allocation_type="本引当",
                    quantity=spec["quantity"],
                    ems_item=ems_item,
                )
            )

        staging = JapanInventoryStaging.query.filter_by(ems_item_id=ems_item.id).first()
        if staging is None and ems.arrived_at is not None:
            db.session.add(
                JapanInventoryStaging(
                    ems_item=ems_item,
                    product_code=ems_item.product_code,
                    product_sub_code=ems_item.product_sub_code,
                    quantity=ems_item.quantity,
                    status="waiting",
                )
            )
            created += 1

    return created


def import_api_data(kind: str) -> tuple[int, str]:
    kind = normalize_text(kind) or ""
    labels = {
        "orders": "Yahoo受注",
        "inventory": "Yahoo在庫",
        "purchases": "発注",
        "ems": "EMS",
        "daniel_purchases": "ダニエル発注",
        "daniel_ems": "ダニエルEMS",
        "tegu_purchases": "テグ発注",
        "tegu_ems": "テグEMS",
        "all": "全データ",
    }
    if kind not in labels:
        raise ValueError("対応していない取込種別です。")

    created = 0
    logical_files: list[tuple[str, str, str]] = []
    today = date.today()

    if kind in {"inventory", "all"}:
        logical_files.append((f"yahoo_inventory_{today:%Y%m%d}.csv", "yahoo_inventory", "inventory"))
    if kind in {"orders", "all"}:
        logical_files.append((f"yahoo_orders_{today:%Y%m%d}.csv", "yahoo_orders", "orders"))
    if kind in {"purchases", "daniel_purchases", "all"}:
        logical_files.append((f"Watanabe_list_{today:%y%m%d}.xlsm", "daniel_purchase", "daniel_purchases"))
    if kind in {"ems", "daniel_ems", "all"}:
        logical_files.append((f"EMS発送リスト_{today:%y%m%d}.xlsx", "daniel_ems", "daniel_ems"))
    if kind in {"purchases", "tegu_purchases", "all"}:
        logical_files.append((f"tegu_purchase_sheet_{today:%Y%m%d}.csv", "tegu_purchase", "tegu_purchases"))
    if kind in {"ems", "tegu_ems", "all"}:
        logical_files.append((f"tegu_ems_sheet_{today:%Y%m%d}.csv", "tegu_ems", "tegu_ems"))

    processed_files = 0
    _ensure_import_inventory()
    for file_name, file_type, operation in logical_files:
        if ImportedFile.query.filter_by(file_name=file_name).first():
            continue

        if operation == "inventory":
            _ensure_import_inventory()
            created_count = Inventory.query.count()
        elif operation == "orders":
            before_count = Order.query.count()
            _ensure_import_orders()
            created_count = max(Order.query.count() - before_count, 0)
        elif operation == "daniel_purchases":
            before_count = Purchase.query.filter_by(source_type="daniel").count()
            _ensure_import_purchases("daniel")
            created_count = max(Purchase.query.filter_by(source_type="daniel").count() - before_count, 0)
        elif operation == "tegu_purchases":
            before_count = Purchase.query.filter_by(source_type="tegu").count()
            _ensure_import_purchases("tegu")
            created_count = max(Purchase.query.filter_by(source_type="tegu").count() - before_count, 0)
        elif operation == "daniel_ems":
            before_count = Ems.query.filter_by(source_type="daniel").count()
            _ensure_import_ems("daniel")
            created_count = max(Ems.query.filter_by(source_type="daniel").count() - before_count, 0)
        elif operation == "tegu_ems":
            before_count = Ems.query.filter_by(source_type="tegu").count()
            _ensure_import_ems("tegu")
            created_count = max(Ems.query.filter_by(source_type="tegu").count() - before_count, 0)
        else:
            created_count = 0

        if _record_imported_file(file_name=file_name, file_type=file_type, record_count=created_count):
            processed_files += 1
            created += created_count

    recalculate_allocations(commit=False)
    db.session.commit()
    run_checks()
    if processed_files == 0:
        return 0, labels[kind]
    return created, labels[kind]


def export_rows(kind: str) -> tuple[str, list[dict]]:
    if kind == "orders":
        rows = [
            {
                "受注番号": order.yahoo_order_id,
                "注文日時": order.ordered_at.isoformat(sep=" ", timespec="minutes"),
                "お客様番号": order.customer_code or "",
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
                "データ元": SOURCE_TYPE_LABELS.get(ems.source_type, ems.source_type),
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
    elif kind == "purchases":
        rows = [
            {
                "データ元": SOURCE_TYPE_LABELS.get(purchase.source_type, purchase.source_type),
                "受注番号": purchase.order_item.order.yahoo_order_id,
                "商品コード": purchase.product_code,
                "数量": purchase.quantity,
                "発注先": purchase.shop_name or "",
                "発注日": purchase.ordered_at.isoformat(),
                "ステータス": purchase.status,
            }
            for purchase in Purchase.query.order_by(Purchase.ordered_at.desc(), Purchase.id.desc()).all()
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
            "imported_files": ImportedFile.query.count(),
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
