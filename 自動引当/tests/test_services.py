from __future__ import annotations

import unittest
from datetime import date, datetime, timedelta

from inventory_tool import create_app
from inventory_tool.config import Config
from inventory_tool.extensions import db
from inventory_tool.models import ImportedFile, Inventory, JapanInventoryStaging, Order, OrderItem
from inventory_tool.services import (
    recalculate_allocations,
    reflect_japan_stock,
    run_checks,
    search_orders_data,
)


class TestConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"


class ServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = create_app(TestConfig)
        self.ctx = self.app.app_context()
        self.ctx.push()
        db.drop_all()
        db.create_all()

    def tearDown(self) -> None:
        db.session.remove()
        db.drop_all()
        self.ctx.pop()

    def test_immediate_inventory_is_reserved_first(self) -> None:
        db.session.add(Inventory(product_code="SKU-A", inventory_type="即納", quantity=3, reserved_qty=0, available_qty=3))
        order = Order(
            yahoo_order_id="TEST-001",
            ordered_at=datetime.now(),
            desired_delivery_date=date.today() + timedelta(days=5),
            customer_name="Tester",
        )
        order.items = [OrderItem(product_code="SKU-A", product_sub_code=None, quantity=2, inventory_type="お取り寄せ")]
        db.session.add(order)
        db.session.commit()

        recalculate_allocations()

        item = OrderItem.query.first()
        inventory = Inventory.query.first()
        self.assertEqual(item.status, "allocated_sokunou")
        self.assertEqual(item.allocated_qty, 2)
        self.assertEqual(inventory.reserved_qty, 2)
        self.assertEqual(inventory.available_qty, 1)

    def test_shortage_creates_alert(self) -> None:
        order = Order(
            yahoo_order_id="TEST-002",
            ordered_at=datetime.now() - timedelta(days=7),
            desired_delivery_date=date.today() + timedelta(days=2),
            customer_name="Tester",
        )
        order.items = [OrderItem(product_code="SKU-B", product_sub_code=None, quantity=1, inventory_type="お取り寄せ")]
        db.session.add(order)
        db.session.commit()

        recalculate_allocations()
        run_checks()

        alerts_page = self.app.test_client().get("/alerts")
        self.assertEqual(alerts_page.status_code, 200)
        self.assertIn("在庫不足", alerts_page.get_data(as_text=True))

    def test_japan_stock_reflection_updates_inventory(self) -> None:
        order = Order(
            yahoo_order_id="TEST-003",
            ordered_at=datetime.now(),
            desired_delivery_date=date.today() + timedelta(days=6),
            customer_name="Tester",
        )
        item = OrderItem(product_code="SKU-C", product_sub_code=None, quantity=1, inventory_type="お取り寄せ")
        order.items = [item]
        db.session.add(order)
        db.session.commit()

        row = JapanInventoryStaging(
            ems_item_id=1,
            product_code="SKU-C",
            product_sub_code=None,
            quantity=2,
            status="to_japan_stock",
        )
        db.session.add(row)
        db.session.commit()

        reflected = reflect_japan_stock()

        inventory = Inventory.query.filter_by(product_code="SKU-C", inventory_type="即納").first()
        self.assertEqual(reflected, 1)
        self.assertIsNotNone(inventory)
        self.assertEqual(inventory.quantity, 2)
        self.assertEqual(row.status, "reflected")

    def test_search_orders_filters_by_product_code(self) -> None:
        order1 = Order(
            yahoo_order_id="TEST-101",
            ordered_at=datetime.now(),
            desired_delivery_date=date.today() + timedelta(days=5),
            customer_code="CUS-101",
            customer_name="Tester A",
        )
        order1.items = [OrderItem(product_code="SKU-SEARCH", product_sub_code=None, quantity=1, inventory_type="即納")]
        order2 = Order(
            yahoo_order_id="TEST-102",
            ordered_at=datetime.now(),
            desired_delivery_date=date.today() + timedelta(days=5),
            customer_code="CUS-102",
            customer_name="Tester B",
        )
        order2.items = [OrderItem(product_code="SKU-OTHER", product_sub_code=None, quantity=1, inventory_type="即納")]
        db.session.add_all([order1, order2])
        db.session.commit()

        results = search_orders_data(product_keyword="SKU-SEARCH")

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].yahoo_order_id, "TEST-101")

    def test_import_orders_route_adds_demo_records(self) -> None:
        response = self.app.test_client().post("/imports/orders", follow_redirects=True)

        self.assertEqual(response.status_code, 200)
        self.assertGreater(Order.query.count(), 0)
        self.assertIn("受注データ", response.get_data(as_text=True))

    def test_import_all_records_imported_files(self) -> None:
        response = self.app.test_client().post("/imports/all", follow_redirects=True)

        self.assertEqual(response.status_code, 200)
        self.assertGreater(ImportedFile.query.count(), 0)

    def test_inline_update_route_updates_order_item_status(self) -> None:
        order = Order(
            yahoo_order_id="TEST-200",
            ordered_at=datetime.now(),
            desired_delivery_date=date.today() + timedelta(days=5),
            customer_code="CUS-200",
            customer_name="Tester Update",
        )
        item = OrderItem(product_code="SKU-UPD", product_sub_code=None, quantity=1, inventory_type="お取り寄せ")
        order.items = [item]
        db.session.add(order)
        db.session.commit()

        response = self.app.test_client().post(
            "/api/update-field",
            json={
                "entity": "order_item",
                "id": item.id,
                "field": "status",
                "value": "fully_allocated",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(OrderItem.query.get(item.id).status, "fully_allocated")


if __name__ == "__main__":
    unittest.main()
