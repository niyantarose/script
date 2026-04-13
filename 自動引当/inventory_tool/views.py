from __future__ import annotations

from datetime import date, datetime

from flask import (
    Blueprint,
    Response,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)

from .constants import (
    ALERT_TYPE_LABELS,
    DELAY_LEVEL_LABELS,
    INVENTORY_TYPE_LABELS,
    JAPAN_STAGING_STATUS_LABELS,
    ORDER_ITEM_STATUS_LABELS,
    ORDER_STATUS_LABELS,
)
from .models import Alert, Ems, Inventory, JapanInventoryStaging, Order, OrderItem, Purchase
from .sample_data import seed_demo_data
from .services import (
    create_ems,
    create_purchase,
    dashboard_stats,
    delay_level_for_order,
    delivery_judgement,
    export_csv,
    latest_orders,
    manual_allocate,
    mark_ems_arrived,
    recalculate_allocations,
    reflect_japan_stock,
    resolve_alert,
    run_checks,
    ship_order_item,
    update_order_meta,
    update_staging_row,
)


main_bp = Blueprint("main", __name__)


def register_template_helpers(app) -> None:
    @app.context_processor
    def inject_labels():
        return {
            "order_item_status_labels": ORDER_ITEM_STATUS_LABELS,
            "order_status_labels": ORDER_STATUS_LABELS,
            "alert_type_labels": ALERT_TYPE_LABELS,
            "delay_level_labels": DELAY_LEVEL_LABELS,
            "japan_staging_status_labels": JAPAN_STAGING_STATUS_LABELS,
            "inventory_type_labels": INVENTORY_TYPE_LABELS,
            "delay_level_for_order": delay_level_for_order,
            "delivery_judgement": delivery_judgement,
        }


@main_bp.get("/")
def dashboard():
    stats = dashboard_stats()
    orders = latest_orders()
    unresolved_alerts = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).limit(10).all()
    inventory_summary = Inventory.query.order_by(Inventory.inventory_type, Inventory.product_code).limit(8).all()
    return render_template(
        "dashboard.html",
        stats=stats,
        orders=orders,
        unresolved_alerts=unresolved_alerts,
        inventory_summary=inventory_summary,
    )


@main_bp.post("/actions/seed-demo")
def seed_demo():
    seed_demo_data()
    flash("デモデータを投入しました。", "success")
    return redirect(url_for("main.dashboard"))


@main_bp.post("/actions/run-checks")
def trigger_checks():
    run_checks()
    flash("4段階チェックと遅延判定を更新しました。", "success")
    return redirect(request.referrer or url_for("main.dashboard"))


@main_bp.get("/orders")
def orders():
    orders_data = Order.query.order_by(Order.ordered_at.desc()).all()
    return render_template("orders.html", orders=orders_data)


@main_bp.post("/orders/<int:order_id>/update")
def update_order(order_id: int):
    update_order_meta(
        order_id=order_id,
        priority_ship_flag=bool(request.form.get("priority_ship_flag")),
        delay_memo=request.form.get("delay_memo"),
        customer_contacted_flag=bool(request.form.get("customer_contacted_flag")),
    )
    flash("受注メモを更新しました。", "success")
    return redirect(url_for("main.orders"))


@main_bp.post("/order-items/<int:order_item_id>/ship")
def ship_item(order_item_id: int):
    ship_order_item(order_item_id)
    flash("発送完了として更新しました。", "success")
    return redirect(url_for("main.orders"))


@main_bp.get("/ems")
def ems():
    ems_list = Ems.query.order_by(Ems.shipped_at.desc()).all()
    return render_template("ems.html", ems_list=ems_list, order_items=OrderItem.query.order_by(OrderItem.id).all())


@main_bp.post("/ems/create")
def create_ems_route():
    try:
        create_ems(
            ems_number=request.form["ems_number"],
            shipped_at=date.fromisoformat(request.form["shipped_at"]),
            memo=request.form.get("memo"),
            items_text=request.form.get("items_text", ""),
        )
        flash("EMS便を登録し、本引当を実行しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
    return redirect(url_for("main.ems"))


@main_bp.post("/ems/<int:ems_id>/arrive")
def arrive_ems_route(ems_id: int):
    try:
        mark_ems_arrived(
            ems_id=ems_id,
            arrived_at=date.fromisoformat(request.form["arrived_at"]),
        )
        flash("EMS入荷を登録し、日本在庫仕分けへ追加しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
    return redirect(url_for("main.ems"))


@main_bp.get("/checks")
def checks():
    alerts = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).all()
    grouped = {}
    for alert_type in ALERT_TYPE_LABELS:
        grouped[alert_type] = [alert for alert in alerts if alert.alert_type == alert_type]
    return render_template("checks.html", grouped_alerts=grouped)


@main_bp.get("/allocation")
def allocation():
    items = OrderItem.query.order_by(OrderItem.id.asc()).all()
    purchases = Purchase.query.order_by(Purchase.ordered_at.desc(), Purchase.id.desc()).all()
    return render_template("allocation.html", items=items, purchases=purchases)


@main_bp.post("/allocation/recalculate")
def recalculate_route():
    recalculate_allocations()
    flash("引当を再計算しました。", "success")
    return redirect(url_for("main.allocation"))


@main_bp.post("/allocation/purchase")
def create_purchase_route():
    try:
        create_purchase(
            order_item_id=int(request.form["order_item_id"]),
            quantity=int(request.form["quantity"]),
            shop_name=request.form.get("shop_name"),
            ordered_at=date.fromisoformat(request.form["ordered_at"]),
            status=request.form["status"],
            memo=request.form.get("memo"),
        )
        flash("発注を登録し、仮引当を記録しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
    return redirect(url_for("main.allocation"))


@main_bp.post("/allocation/manual")
def manual_allocation_route():
    try:
        manual_allocate(
            order_item_id=int(request.form["order_item_id"]),
            inventory_type=request.form["inventory_type"],
            quantity=int(request.form["quantity"]),
            allocated_by=request.form.get("allocated_by"),
        )
        flash("手動引当を実行しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
    return redirect(url_for("main.allocation"))


@main_bp.get("/japan-stock")
def japan_stock():
    rows = JapanInventoryStaging.query.order_by(JapanInventoryStaging.created_at.desc()).all()
    return render_template("japan_stock.html", rows=rows, order_items=OrderItem.query.order_by(OrderItem.id).all())


@main_bp.post("/japan-stock/<int:staging_id>/update")
def japan_stock_update(staging_id: int):
    try:
        update_staging_row(
            staging_id=staging_id,
            action=request.form["action"],
            assigned_order_item_id=int(request.form["assigned_order_item_id"])
            if request.form.get("assigned_order_item_id")
            else None,
            excluded_reason=request.form.get("excluded_reason"),
        )
        flash("日本在庫仕分けを更新しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
    return redirect(url_for("main.japan_stock"))


@main_bp.post("/japan-stock/reflect")
def japan_stock_reflect():
    count = reflect_japan_stock()
    flash(f"{count}件を日本在庫へ反映しました。", "success")
    return redirect(url_for("main.japan_stock"))


@main_bp.get("/alerts")
def alerts():
    unresolved = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).all()
    resolved = Alert.query.filter_by(resolved_flag=True).order_by(Alert.created_at.desc()).all()
    return render_template("alerts.html", unresolved=unresolved, resolved=resolved)


@main_bp.post("/alerts/<int:alert_id>/resolve")
def resolve_alert_route(alert_id: int):
    resolve_alert(alert_id)
    flash("アラートを解決済みにしました。", "success")
    return redirect(url_for("main.alerts"))


@main_bp.get("/export/<kind>.csv")
def export(kind: str):
    try:
        csv_text = export_csv(kind)
    except Exception as exc:  # noqa: BLE001
        flash(str(exc), "error")
        return redirect(request.referrer or url_for("main.dashboard"))

    return Response(
        csv_text,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={kind}_{datetime.now():%Y%m%d_%H%M}.csv"},
    )
