from __future__ import annotations

from datetime import date, datetime

from flask import (
    Blueprint,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)
from sqlalchemy.exc import IntegrityError

from .constants import (
    ALERT_TYPE_LABELS,
    DELAY_LEVEL_LABELS,
    EMS_STATUS_LABELS,
    INVENTORY_TYPE_LABELS,
    JAPAN_STAGING_STATUS_LABELS,
    ORDER_ITEM_STATUS_LABELS,
    ORDER_STATUS_LABELS,
    PURCHASE_STATUS_LABELS,
    SOURCE_TYPE_LABELS,
)
from .extensions import db
from .models import Alert, Ems, Inventory, JapanInventoryStaging, Order, OrderItem, Purchase
from .sample_data import seed_demo_data
from .services import (
    create_ems,
    create_purchase,
    dashboard_stats,
    delay_level_for_order,
    delivery_status_for_item,
    delivery_judgement,
    export_csv,
    import_api_data,
    latest_ems,
    latest_orders,
    latest_purchases,
    latest_arrived_ems_number,
    latest_ems_for_item,
    manual_allocate,
    mark_ems_arrived,
    order_status_counts,
    purchase_status_for_item,
    recalculate_allocations,
    reflect_japan_stock,
    resolve_alert,
    run_checks,
    search_orders_data,
    ship_order_item,
    update_record_field,
    update_order_meta,
    update_staging_row,
)


main_bp = Blueprint("main", __name__)


def format_error_message(exc: Exception) -> str:
    raw = str(exc).strip()
    if not raw:
        return "処理に失敗しました。入力内容を確認してください。"
    if "Invalid isoformat string" in raw or "invalid literal for int()" in raw:
        return "入力形式が正しくありません。日付や数値を確認してください。"
    if "404 Not Found" in raw or "Not Found" in raw:
        return "対象データが見つかりませんでした。"
    return raw


def register_template_helpers(app) -> None:
    @app.context_processor
    def inject_labels():
        return {
            "order_item_status_labels": ORDER_ITEM_STATUS_LABELS,
            "order_status_labels": ORDER_STATUS_LABELS,
            "purchase_status_labels": PURCHASE_STATUS_LABELS,
            "ems_status_labels": EMS_STATUS_LABELS,
            "alert_type_labels": ALERT_TYPE_LABELS,
            "delay_level_labels": DELAY_LEVEL_LABELS,
            "japan_staging_status_labels": JAPAN_STAGING_STATUS_LABELS,
            "inventory_type_labels": INVENTORY_TYPE_LABELS,
            "source_type_labels": SOURCE_TYPE_LABELS,
            "delay_level_for_order": delay_level_for_order,
            "delivery_status_for_item": delivery_status_for_item,
            "delivery_judgement": delivery_judgement,
            "latest_arrived_ems_number": latest_arrived_ems_number,
            "latest_ems_for_item": latest_ems_for_item,
            "purchase_status_for_item": purchase_status_for_item,
        }


def _date_text(value: date | None) -> str:
    return value.isoformat() if value else ""


def _datetime_text(value: datetime | None) -> str:
    return value.strftime("%Y-%m-%d %H:%M") if value else ""


def _alert_level(alert_type: str) -> str:
    if alert_type in {"delay_warning", "stock_shortage", "purchase_missing"}:
        return "danger"
    if alert_type in {"korea_ship_missing", "japan_arrival_missing"}:
        return "warning"
    return "notice"


def _latest_purchase_for_source(order_item: OrderItem, source_type: str) -> Purchase | None:
    rows = [purchase for purchase in order_item.purchases if purchase.source_type == source_type]
    if not rows:
        return None
    return max(rows, key=lambda purchase: (purchase.ordered_at or date.min, purchase.id or 0))


def _latest_ems_for_source(order_item: OrderItem, source_type: str) -> Ems | None:
    rows = [
        ems_item.ems
        for ems_item in order_item.ems_items
        if ems_item.ems is not None and ems_item.ems.source_type == source_type
    ]
    if not rows:
        return None
    return max(rows, key=lambda ems: (ems.shipped_at or date.min, ems.id or 0))


def _latest_arrived_ems_for_source(order_item: OrderItem, source_type: str) -> str:
    rows = [
        ems_item.ems
        for ems_item in order_item.ems_items
        if ems_item.ems is not None and ems_item.ems.source_type == source_type and ems_item.ems.arrived_at is not None
    ]
    if not rows:
        return ""
    latest = max(rows, key=lambda ems: (ems.arrived_at or date.min, ems.id or 0))
    return latest.ems_number


def _serialize_order_rows(orders: list[Order]) -> list[dict]:
    rows: list[dict] = []
    for order in orders:
        delay_level = delay_level_for_order(order)
        item_count = len(order.items)
        for line_number, item in enumerate(order.items, start=1):
            rows.append(
                {
                    "row_id": f"order-item-{item.id}",
                    "group_key": order.yahoo_order_id,
                    "order_id": order.id,
                    "order_item_id": item.id,
                    "order_number": order.yahoo_order_id,
                    "line_number": line_number,
                    "item_count": item_count,
                    "product_code": item.product_code,
                    "delivery_status": delivery_status_for_item(item),
                    "item_status": item.status,
                    "item_status_label": ORDER_ITEM_STATUS_LABELS.get(item.status, item.status),
                    "arrived_ems_number": latest_arrived_ems_number(item),
                    "delay_level": delay_level,
                    "desired_delivery_date": _date_text(order.desired_delivery_date),
                    "customer_code": order.customer_code or "",
                    "customer_name": order.customer_name or "",
                }
            )
    return rows


def _serialize_search_rows(orders: list[Order]) -> list[dict]:
    rows: list[dict] = []
    for order in orders:
        delay_level = delay_level_for_order(order)
        for line_number, item in enumerate(order.items, start=1):
            rows.append(
                {
                    "row_id": f"search-item-{item.id}",
                    "group_key": order.yahoo_order_id,
                    "order_id": order.id,
                    "order_item_id": item.id,
                    "order_number": order.yahoo_order_id,
                    "ordered_at": _datetime_text(order.ordered_at),
                    "customer_code": order.customer_code or "",
                    "customer_name": order.customer_name or "",
                    "desired_delivery_date": _date_text(order.desired_delivery_date),
                    "order_status": order.status,
                    "order_status_label": ORDER_STATUS_LABELS.get(order.status, order.status),
                    "delay_level": delay_level,
                    "delay_label": DELAY_LEVEL_LABELS.get(delay_level, delay_level),
                    "line_number": line_number,
                    "product_code": item.product_code,
                    "product_sub_code": item.product_sub_code or "",
                    "quantity": item.quantity,
                    "inventory_type": item.inventory_type,
                    "item_status": item.status,
                    "item_status_label": ORDER_ITEM_STATUS_LABELS.get(item.status, item.status),
                }
            )
    return rows


def _serialize_purchase_rows(orders: list[Order], source_type: str) -> list[dict]:
    rows: list[dict] = []
    for order in orders:
        for line_number, item in enumerate(order.items, start=1):
            purchase = _latest_purchase_for_source(item, source_type)
            rows.append(
                {
                    "row_id": f"purchase-{source_type}-{item.id}",
                    "group_key": order.yahoo_order_id,
                    "order_id": order.id,
                    "order_item_id": item.id,
                    "purchase_id": purchase.id if purchase else None,
                    "order_number": order.yahoo_order_id,
                    "line_number": line_number,
                    "product_code": purchase.product_code if purchase else item.product_code,
                    "shop_name": purchase.shop_name or "" if purchase else "",
                    "ordered_at": _date_text(purchase.ordered_at) if purchase else "",
                    "purchase_status": purchase.status if purchase else "",
                    "purchase_status_label": PURCHASE_STATUS_LABELS.get(purchase.status, purchase.status)
                    if purchase
                    else purchase_status_for_item(item),
                    "allocation_status": item.status,
                    "allocation_status_label": ORDER_ITEM_STATUS_LABELS.get(item.status, item.status),
                    "memo": purchase.memo or "" if purchase else "",
                    "has_purchase": purchase is not None,
                }
            )
    return rows


def _serialize_ems_rows(orders: list[Order], source_type: str) -> list[dict]:
    rows: list[dict] = []
    for order in orders:
        for line_number, item in enumerate(order.items, start=1):
            ems = _latest_ems_for_source(item, source_type)
            rows.append(
                {
                    "row_id": f"ems-{source_type}-{item.id}",
                    "group_key": order.yahoo_order_id,
                    "order_id": order.id,
                    "order_item_id": item.id,
                    "ems_id": ems.id if ems else None,
                    "order_number": order.yahoo_order_id,
                    "line_number": line_number,
                    "product_code": item.product_code,
                    "quantity": item.quantity,
                    "ems_number": ems.ems_number if ems else "",
                    "shipped_at": _date_text(ems.shipped_at) if ems else "",
                    "estimated_arrival": _date_text(ems.estimated_arrival) if ems else "",
                    "arrived_at": _date_text(ems.arrived_at) if ems and ems.arrived_at else "",
                    "ems_status": ems.status if ems else "",
                    "ems_status_label": EMS_STATUS_LABELS.get(ems.status, ems.status) if ems else "未登録",
                    "memo": ems.memo or "" if ems else "",
                    "has_ems": ems is not None,
                    "arrived_ems_number": _latest_arrived_ems_for_source(item, source_type),
                }
            )
    return rows


def _serialize_japan_stock_rows(rows: list[JapanInventoryStaging]) -> list[dict]:
    return [
        {
            "row_id": f"japan-{row.id}",
            "staging_id": row.id,
            "ems_number": row.ems_item.ems.ems_number,
            "arrived_at": _date_text(row.ems_item.ems.arrived_at),
            "product_code": row.product_code,
            "quantity": row.quantity,
            "status": row.status,
            "status_label": JAPAN_STAGING_STATUS_LABELS.get(row.status, row.status),
            "assigned_order_item_id": row.assigned_order_item_id or "",
            "excluded_reason": row.excluded_reason or "",
        }
        for row in rows
    ]


def _serialize_alert_rows(alerts: list[Alert]) -> list[dict]:
    return [
        {
            "row_id": f"alert-{alert.id}",
            "alert_id": alert.id,
            "alert_type": alert.alert_type,
            "alert_type_label": ALERT_TYPE_LABELS.get(alert.alert_type, alert.alert_type),
            "order_number": alert.order.yahoo_order_id if alert.order else "共通",
            "product_code": alert.product_code or "",
            "message": alert.message,
            "created_at": _datetime_text(alert.created_at),
            "level": _alert_level(alert.alert_type),
            "resolved_flag": alert.resolved_flag,
        }
        for alert in alerts
    ]


def _serialize_purchase_feed_rows(rows: list[Purchase]) -> list[dict]:
    return [
        {
            "row_id": f"purchase-feed-{row.id}",
            "order_number": row.order_item.order.yahoo_order_id,
            "product_code": row.product_code,
            "shop_name": row.shop_name or "",
            "ordered_at": _date_text(row.ordered_at),
            "status_label": PURCHASE_STATUS_LABELS.get(row.status, row.status),
        }
        for row in rows
    ]


def _serialize_ems_feed_rows(rows: list[Ems]) -> list[dict]:
    return [
        {
            "row_id": f"ems-feed-{row.id}",
            "ems_number": row.ems_number,
            "shipped_at": _date_text(row.shipped_at),
            "estimated_arrival": _date_text(row.estimated_arrival),
            "arrived_at": _date_text(row.arrived_at),
            "status_label": EMS_STATUS_LABELS.get(row.status, row.status),
            "memo": row.memo or "",
        }
        for row in rows
    ]


def _serialize_dashboard_order_rows(orders: list[Order]) -> list[dict]:
    return [
        {
            "row_id": f"dashboard-order-{order.id}",
            "order_number": order.yahoo_order_id,
            "customer_name": order.customer_name or "",
            "desired_delivery_date": _date_text(order.desired_delivery_date),
            "status_label": ORDER_STATUS_LABELS.get(order.status, order.status),
            "delay_level": delay_level_for_order(order),
            "delay_label": DELAY_LEVEL_LABELS.get(delay_level_for_order(order), ""),
        }
        for order in orders
    ]


def _serialize_inventory_rows(rows: list[Inventory]) -> list[dict]:
    return [
        {
            "row_id": f"inventory-{row.id}",
            "product_code": row.product_code,
            "product_sub_code": row.product_sub_code or "",
            "inventory_type": row.inventory_type,
            "quantity": row.quantity,
            "reserved_qty": row.reserved_qty,
            "available_qty": row.available_qty,
        }
        for row in rows
    ]


@main_bp.get("/")
def dashboard():
    stats = dashboard_stats()
    orders = latest_orders()
    unresolved_alerts = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).limit(10).all()
    inventory_summary = Inventory.query.order_by(Inventory.inventory_type, Inventory.product_code).limit(8).all()
    purchase_feeds = {
        "daniel": latest_purchases("daniel"),
        "tegu": latest_purchases("tegu"),
    }
    ems_feeds = {
        "daniel": latest_ems("daniel"),
        "tegu": latest_ems("tegu"),
    }
    return render_template(
        "dashboard.html",
        stats=stats,
        unresolved_alerts=unresolved_alerts,
        dashboard_order_rows=_serialize_dashboard_order_rows(orders),
        inventory_rows=_serialize_inventory_rows(inventory_summary),
        purchase_feed_rows={
            "daniel": _serialize_purchase_feed_rows(purchase_feeds["daniel"]),
            "tegu": _serialize_purchase_feed_rows(purchase_feeds["tegu"]),
        },
        ems_feed_rows={
            "daniel": _serialize_ems_feed_rows(ems_feeds["daniel"]),
            "tegu": _serialize_ems_feed_rows(ems_feeds["tegu"]),
        },
        alert_rows=_serialize_alert_rows(unresolved_alerts),
        order_status_summary=order_status_counts(),
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


@main_bp.post("/imports/<kind>")
def import_data_route(kind: str):
    try:
        created, label = import_api_data(kind)
        if created > 0:
            flash(f"{label}データを {created} 件取り込みました。", "success")
        else:
            flash(
                f"{label}データの新規取込はありませんでした。現在の画面確認用データをそのまま使用しています。",
                "success",
            )
    except Exception as exc:  # noqa: BLE001
        flash(format_error_message(exc), "error")
    return redirect(request.referrer or url_for("main.dashboard"))


@main_bp.post("/api/update-field")
def update_field_route():
    payload = request.get_json(silent=True) or request.form
    try:
        record = update_record_field(
            entity=payload.get("entity"),
            record_id=int(payload.get("id")),
            field_name=payload.get("field"),
            value=payload.get("value"),
            edited_by="web",
        )
        return jsonify({"ok": True, "record_id": record.id})
    except IntegrityError:
        db.session.rollback()
        return jsonify({"ok": False, "message": "同じ値がすでに登録されています。"}), 400
    except Exception as exc:  # noqa: BLE001
        db.session.rollback()
        return jsonify({"ok": False, "message": format_error_message(exc)}), 400


@main_bp.get("/orders")
def orders():
    orders_data = Order.query.order_by(Order.ordered_at.desc()).all()
    return render_template("orders.html", order_rows=_serialize_order_rows(orders_data))


@main_bp.get("/order-search")
def order_search():
    order_keyword = request.args.get("order_keyword", "")
    customer_keyword = request.args.get("customer_keyword", "")
    product_keyword = request.args.get("product_keyword", "")
    results = search_orders_data(
        order_keyword=order_keyword,
        customer_keyword=customer_keyword,
        product_keyword=product_keyword,
    )
    return render_template(
        "order_search.html",
        search_rows=_serialize_search_rows(results),
        order_keyword=order_keyword,
        customer_keyword=customer_keyword,
        product_keyword=product_keyword,
        has_filters=bool(order_keyword or customer_keyword or product_keyword),
    )


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


def _purchase_page(source_type: str):
    orders_data = Order.query.order_by(Order.ordered_at.desc()).all()
    return render_template(
        "allocation.html",
        purchase_rows=_serialize_purchase_rows(orders_data, source_type),
        source_type=source_type,
        page_title=f"{SOURCE_TYPE_LABELS[source_type]}発注リスト",
        import_kind=f"{source_type}_purchases",
    )


def _ems_page(source_type: str):
    orders_data = Order.query.order_by(Order.ordered_at.desc()).all()
    return render_template(
        "ems.html",
        ems_rows=_serialize_ems_rows(orders_data, source_type),
        source_type=source_type,
        page_title=f"{SOURCE_TYPE_LABELS[source_type]}EMSリスト",
        import_kind=f"{source_type}_ems",
    )


@main_bp.get("/daniel-purchases")
def daniel_purchases():
    return _purchase_page("daniel")


@main_bp.get("/tegu-purchases")
def tegu_purchases():
    return _purchase_page("tegu")


@main_bp.get("/daniel-ems")
def daniel_ems():
    return _ems_page("daniel")


@main_bp.get("/tegu-ems")
def tegu_ems():
    return _ems_page("tegu")


@main_bp.get("/ems")
def ems():
    return redirect(url_for("main.daniel_ems"))


@main_bp.post("/ems/create")
def create_ems_route():
    try:
        source_type = request.form.get("source_type", "daniel")
        create_ems(
            ems_number=request.form["ems_number"],
            shipped_at=date.fromisoformat(request.form["shipped_at"]),
            memo=request.form.get("memo"),
            items_text=request.form.get("items_text", ""),
            source_type=source_type,
        )
        flash("EMS便を登録し、本引当を実行しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(format_error_message(exc), "error")
    endpoint = "main.tegu_ems" if request.form.get("source_type") == "tegu" else "main.daniel_ems"
    return redirect(url_for(endpoint))


@main_bp.post("/ems/<int:ems_id>/arrive")
def arrive_ems_route(ems_id: int):
    try:
        mark_ems_arrived(
            ems_id=ems_id,
            arrived_at=date.fromisoformat(request.form["arrived_at"]),
        )
        flash("EMS入荷を登録し、日本在庫仕分けへ追加しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(format_error_message(exc), "error")
    referrer = request.referrer or ""
    if "tegu-ems" in referrer:
        return redirect(url_for("main.tegu_ems"))
    return redirect(url_for("main.daniel_ems"))


@main_bp.get("/checks")
def checks():
    alerts = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).all()
    grouped = {}
    for alert_type in ALERT_TYPE_LABELS:
        grouped[alert_type] = [alert for alert in alerts if alert.alert_type == alert_type]
    return render_template("checks.html", check_rows=_serialize_alert_rows(alerts), grouped_alerts=grouped)


@main_bp.get("/allocation")
def allocation():
    return redirect(url_for("main.daniel_purchases"))


@main_bp.post("/allocation/recalculate")
def recalculate_route():
    recalculate_allocations()
    flash("引当を再計算しました。", "success")
    return redirect(request.referrer or url_for("main.daniel_purchases"))


@main_bp.post("/allocation/purchase")
def create_purchase_route():
    try:
        source_type = request.form.get("source_type", "daniel")
        create_purchase(
            order_item_id=int(request.form["order_item_id"]),
            quantity=int(request.form["quantity"]),
            shop_name=request.form.get("shop_name"),
            ordered_at=date.fromisoformat(request.form["ordered_at"]),
            status=request.form["status"],
            memo=request.form.get("memo"),
            source_type=source_type,
        )
        flash("発注を登録し、仮引当を記録しました。", "success")
    except Exception as exc:  # noqa: BLE001
        flash(format_error_message(exc), "error")
    endpoint = "main.tegu_purchases" if request.form.get("source_type") == "tegu" else "main.daniel_purchases"
    return redirect(url_for(endpoint))


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
        flash(format_error_message(exc), "error")
    return redirect(request.referrer or url_for("main.daniel_purchases"))


@main_bp.get("/japan-stock")
def japan_stock():
    rows = JapanInventoryStaging.query.order_by(JapanInventoryStaging.created_at.desc()).all()
    return render_template("japan_stock.html", stock_rows=_serialize_japan_stock_rows(rows))


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
        flash(format_error_message(exc), "error")
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
    return render_template(
        "alerts.html",
        unresolved_rows=_serialize_alert_rows(unresolved),
        resolved_rows=_serialize_alert_rows(resolved),
    )


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
        flash(format_error_message(exc), "error")
        return redirect(request.referrer or url_for("main.dashboard"))

    return Response(
        csv_text,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={kind}_{datetime.now():%Y%m%d_%H%M}.csv"},
    )
