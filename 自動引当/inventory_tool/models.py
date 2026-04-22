from __future__ import annotations

from datetime import date

from sqlalchemy import func
from sqlalchemy.orm import relationship

from .extensions import db


class TimestampMixin:
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )


class Order(TimestampMixin, db.Model):
    __tablename__ = "orders"

    id = db.Column(db.Integer, primary_key=True)
    yahoo_order_id = db.Column(db.String(50), nullable=False, unique=True)
    ordered_at = db.Column(db.DateTime, nullable=False)
    desired_delivery_date = db.Column(db.Date)
    customer_code = db.Column(db.String(100))
    customer_name = db.Column(db.String(100))
    priority_ship_flag = db.Column(db.Boolean, nullable=False, default=False)
    yahoo_ship_status = db.Column(db.String(30))
    status = db.Column(db.String(30), nullable=False, default="pending")
    delay_memo = db.Column(db.Text)
    customer_contacted_flag = db.Column(db.Boolean, nullable=False, default=False)

    items = relationship("OrderItem", back_populates="order", cascade="all, delete-orphan")
    alerts = relationship("Alert", back_populates="order")

    @property
    def ordered_date(self) -> date:
        return self.ordered_at.date()


class OrderItem(TimestampMixin, db.Model):
    __tablename__ = "order_items"

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey("orders.id"), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100))
    quantity = db.Column(db.Integer, nullable=False)
    inventory_type = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(30), nullable=False, default="pending")
    allocated_qty = db.Column(db.Integer, nullable=False, default=0)
    shipped_flag = db.Column(db.Boolean, nullable=False, default=False)

    order = relationship("Order", back_populates="items")
    purchases = relationship("Purchase", back_populates="order_item", cascade="all, delete-orphan")
    ems_items = relationship("EmsItem", back_populates="order_item")
    allocations = relationship("Allocation", back_populates="order_item")
    alerts = relationship("Alert", back_populates="order_item")
    staging_assignments = relationship(
        "JapanInventoryStaging",
        back_populates="assigned_order_item",
        foreign_keys="JapanInventoryStaging.assigned_order_item_id",
    )


class Purchase(TimestampMixin, db.Model):
    __tablename__ = "purchases"

    id = db.Column(db.Integer, primary_key=True)
    order_item_id = db.Column(db.Integer, db.ForeignKey("order_items.id"), nullable=False)
    source_type = db.Column(db.String(20), nullable=False, default="daniel")
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100))
    quantity = db.Column(db.Integer, nullable=False)
    shop_name = db.Column(db.String(100))
    ordered_at = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(30), nullable=False, default="ordered")
    memo = db.Column(db.Text)

    order_item = relationship("OrderItem", back_populates="purchases")


class Ems(TimestampMixin, db.Model):
    __tablename__ = "ems"

    id = db.Column(db.Integer, primary_key=True)
    source_type = db.Column(db.String(20), nullable=False, default="daniel")
    ems_number = db.Column(db.String(50), nullable=False, unique=True)
    shipped_at = db.Column(db.Date, nullable=False)
    estimated_arrival = db.Column(db.Date, nullable=False)
    arrived_at = db.Column(db.Date)
    status = db.Column(db.String(30), nullable=False, default="in_transit")
    memo = db.Column(db.Text)

    items = relationship("EmsItem", back_populates="ems", cascade="all, delete-orphan")


class EmsItem(db.Model):
    __tablename__ = "ems_items"

    id = db.Column(db.Integer, primary_key=True)
    ems_id = db.Column(db.Integer, db.ForeignKey("ems.id"), nullable=False)
    order_item_id = db.Column(db.Integer, db.ForeignKey("order_items.id"), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100))
    quantity = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())

    ems = relationship("Ems", back_populates="items")
    order_item = relationship("OrderItem", back_populates="ems_items")
    staging_rows = relationship("JapanInventoryStaging", back_populates="ems_item")
    allocations = relationship("Allocation", back_populates="ems_item")


class Inventory(db.Model):
    __tablename__ = "inventory"

    id = db.Column(db.Integer, primary_key=True)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100))
    inventory_type = db.Column(db.String(20), nullable=False)
    quantity = db.Column(db.Integer, nullable=False, default=0)
    reserved_qty = db.Column(db.Integer, nullable=False, default=0)
    available_qty = db.Column(db.Integer, nullable=False, default=0)
    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )


class Allocation(db.Model):
    __tablename__ = "allocations"

    id = db.Column(db.Integer, primary_key=True)
    order_item_id = db.Column(db.Integer, db.ForeignKey("order_items.id"), nullable=False)
    inventory_type = db.Column(db.String(20), nullable=False)
    allocation_type = db.Column(db.String(20), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    ems_item_id = db.Column(db.Integer, db.ForeignKey("ems_items.id"))
    allocated_by = db.Column(db.String(50))
    allocated_at = db.Column(db.DateTime, nullable=False, server_default=func.now())

    order_item = relationship("OrderItem", back_populates="allocations")
    ems_item = relationship("EmsItem", back_populates="allocations")


class Alert(db.Model):
    __tablename__ = "alerts"

    id = db.Column(db.Integer, primary_key=True)
    alert_type = db.Column(db.String(50), nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey("orders.id"))
    order_item_id = db.Column(db.Integer, db.ForeignKey("order_items.id"))
    product_code = db.Column(db.String(100))
    message = db.Column(db.Text, nullable=False)
    resolved_flag = db.Column(db.Boolean, nullable=False, default=False)
    resolved_at = db.Column(db.DateTime)
    created_at = db.Column(db.DateTime, nullable=False, server_default=func.now())

    order = relationship("Order", back_populates="alerts")
    order_item = relationship("OrderItem", back_populates="alerts")


class JapanInventoryStaging(TimestampMixin, db.Model):
    __tablename__ = "japan_inventory_staging"

    id = db.Column(db.Integer, primary_key=True)
    ems_item_id = db.Column(db.Integer, db.ForeignKey("ems_items.id"), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100))
    quantity = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(30), nullable=False, default="waiting")
    assigned_order_item_id = db.Column(db.Integer, db.ForeignKey("order_items.id"))
    reflected_at = db.Column(db.DateTime)
    excluded_reason = db.Column(db.Text)

    ems_item = relationship("EmsItem", back_populates="staging_rows")
    assigned_order_item = relationship(
        "OrderItem",
        back_populates="staging_assignments",
        foreign_keys=[assigned_order_item_id],
    )


class ImportedFile(db.Model):
    __tablename__ = "imported_files"

    id = db.Column(db.Integer, primary_key=True)
    file_name = db.Column(db.String(255), nullable=False, unique=True)
    file_type = db.Column(db.String(50), nullable=False)
    imported_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
    record_count = db.Column(db.Integer, nullable=False, default=0)


class EditLog(db.Model):
    __tablename__ = "edit_logs"

    id = db.Column(db.Integer, primary_key=True)
    table_name = db.Column(db.String(50), nullable=False)
    record_id = db.Column(db.Integer, nullable=False)
    field_name = db.Column(db.String(100), nullable=False)
    old_value = db.Column(db.Text)
    new_value = db.Column(db.Text)
    edited_by = db.Column(db.String(100))
    edited_at = db.Column(db.DateTime, nullable=False, server_default=func.now())
