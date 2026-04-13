from __future__ import annotations

from datetime import date, datetime, timedelta

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
from .services import recalculate_allocations, run_checks


def reset_demo_data() -> None:
    for model in [JapanInventoryStaging, Alert, Allocation, EmsItem, Ems, Purchase, OrderItem, Order, Inventory]:
        model.query.delete()
    db.session.commit()


def seed_demo_data() -> None:
    reset_demo_data()

    inventories = [
        Inventory(product_code="SKU-1001", product_sub_code=None, inventory_type="即納", quantity=5, reserved_qty=0, available_qty=5),
        Inventory(product_code="SKU-1002", product_sub_code="BLUE", inventory_type="即納", quantity=1, reserved_qty=0, available_qty=1),
        Inventory(product_code="SKU-2001", product_sub_code=None, inventory_type="お取り寄せ", quantity=99, reserved_qty=0, available_qty=99),
    ]
    db.session.add_all(inventories)

    order1 = Order(
        yahoo_order_id="YH-20260413-001",
        ordered_at=datetime.now() - timedelta(days=1),
        desired_delivery_date=date.today() + timedelta(days=6),
        customer_name="田中 花子",
        status="pending",
    )
    order1.items = [
        OrderItem(product_code="SKU-1001", product_sub_code=None, quantity=2, inventory_type="即納"),
    ]

    order2 = Order(
        yahoo_order_id="YH-20260413-002",
        ordered_at=datetime.now() - timedelta(days=3),
        desired_delivery_date=date.today() + timedelta(days=5),
        customer_name="佐藤 次郎",
        status="pending",
    )
    order2.items = [
        OrderItem(product_code="SKU-1002", product_sub_code="BLUE", quantity=2, inventory_type="お取り寄せ"),
    ]

    order3 = Order(
        yahoo_order_id="YH-20260413-003",
        ordered_at=datetime.now() - timedelta(days=7),
        desired_delivery_date=date.today() + timedelta(days=3),
        customer_name="鈴木 一郎",
        status="pending",
        delay_memo="仕入先確認中",
    )
    order3.items = [
        OrderItem(product_code="SKU-3001", product_sub_code=None, quantity=1, inventory_type="お取り寄せ"),
    ]

    order4 = Order(
        yahoo_order_id="YH-20260413-004",
        ordered_at=datetime.now() - timedelta(days=4),
        desired_delivery_date=date.today() + timedelta(days=7),
        customer_name="高橋 美咲",
        status="pending",
    )
    order4.items = [
        OrderItem(product_code="SKU-2001", product_sub_code=None, quantity=1, inventory_type="お取り寄せ"),
    ]

    db.session.add_all([order1, order2, order3, order4])
    db.session.flush()

    purchase1 = Purchase(
        order_item=order2.items[0],
        product_code="SKU-1002",
        product_sub_code="BLUE",
        quantity=1,
        shop_name="K-Shop",
        ordered_at=date.today() - timedelta(days=1),
        status="arrived",
    )
    purchase2 = Purchase(
        order_item=order4.items[0],
        product_code="SKU-2001",
        product_sub_code=None,
        quantity=1,
        shop_name="K-Mall",
        ordered_at=date.today() - timedelta(days=2),
        status="arrived",
    )
    db.session.add_all([purchase1, purchase2])
    db.session.flush()

    db.session.add_all(
        [
            Allocation(order_item=order2.items[0], inventory_type="お取り寄せ", allocation_type="仮引当", quantity=1),
            Allocation(order_item=order4.items[0], inventory_type="お取り寄せ", allocation_type="仮引当", quantity=1),
        ]
    )

    ems = Ems(
        ems_number="EMS123456789KR",
        shipped_at=date.today() - timedelta(days=4),
        estimated_arrival=date.today() - timedelta(days=2),
        arrived_at=date.today() - timedelta(days=1),
        status="arrived",
        memo="4月便",
    )
    db.session.add(ems)
    db.session.flush()

    ems_item = EmsItem(
        ems=ems,
        order_item=order4.items[0],
        product_code="SKU-2001",
        product_sub_code=None,
        quantity=1,
    )
    db.session.add(ems_item)
    db.session.flush()
    db.session.add(
        Allocation(
            order_item=order4.items[0],
            inventory_type="お取り寄せ",
            allocation_type="本引当",
            quantity=1,
            ems_item=ems_item,
        )
    )
    db.session.add(
        JapanInventoryStaging(
            ems_item=ems_item,
            product_code="SKU-2001",
            product_sub_code=None,
            quantity=1,
            status="waiting",
        )
    )

    recalculate_allocations(commit=False)
    db.session.commit()
    run_checks()
