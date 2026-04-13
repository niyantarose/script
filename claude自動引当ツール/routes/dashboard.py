from flask import Blueprint, render_template
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.alert import Alert
from sqlalchemy import func

bp = Blueprint('dashboard', __name__)


@bp.route('/')
@bp.route('/dashboard')
def index():
    # 発送可能件数: 全商品がfully_allocatedの受注
    shippable = db.session.query(Order).filter(
        ~Order.id.in_(
            db.session.query(OrderItem.order_id).filter(
                OrderItem.status != 'fully_allocated'
            ).distinct()
        ),
        Order.status != 'shipped'
    ).count()

    # 部分在庫待ち
    partial = db.session.query(func.count(func.distinct(OrderItem.order_id))).filter(
        OrderItem.status == 'partial_waiting'
    ).scalar() or 0

    # 遅延中（注文日から7日以上経過で未発送）
    from datetime import datetime, timedelta
    threshold = datetime.now() - timedelta(days=7)
    delayed = db.session.query(Order).filter(
        Order.ordered_at < threshold,
        Order.status != 'shipped'
    ).count()

    # 発注漏れ
    purchase_missing = db.session.query(OrderItem).filter(
        OrderItem.inventory_type == 'お取り寄せ',
        OrderItem.status == 'pending',
        ~OrderItem.id.in_(
            db.session.query(Purchase.order_item_id).distinct()
        )
    ).count()

    # 未解決アラート
    alerts = Alert.query.filter_by(resolved_flag=False).order_by(Alert.created_at.desc()).limit(10).all()

    return render_template('dashboard.html',
                           shippable=shippable,
                           partial=partial,
                           delayed=delayed,
                           purchase_missing=purchase_missing,
                           alerts=alerts)
