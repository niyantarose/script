from flask import Blueprint, render_template, request
from models.order import Order
from models.order_item import OrderItem
from models.ems_item import EmsItem

bp = Blueprint('order_search', __name__, url_prefix='/order-search')


def _build_result(o):
    """注文の全商品を取得してグループ化"""
    items = OrderItem.query.filter_by(order_id=o.id).all()
    ems_numbers = set()
    for item in items:
        for ei in EmsItem.query.filter_by(order_item_id=item.id).all():
            ems_numbers.add(ei.ems.ems_number)
    return {
        'order': o,
        'order_items': items,
        'ems_numbers': '、'.join(sorted(ems_numbers)) if ems_numbers else '',
    }


@bp.route('/')
def index():
    order_no = request.args.get('order_no', '').strip()
    customer = request.args.get('customer', '').strip()
    product_code = request.args.get('product_code', '').strip()

    searched = bool(order_no or customer or product_code)
    results = []

    if searched:
        order_ids = set()

        # 商品コードにヒットした注文のIDを収集
        if product_code:
            matched_items = OrderItem.query.filter(
                OrderItem.product_code.contains(product_code)
            ).all()
            for i in matched_items:
                order_ids.add(i.order_id)

        # 注文番号・お客様名で絞り込む
        if order_no or customer:
            q = Order.query
            if order_no:
                q = q.filter(Order.yahoo_order_id.contains(order_no))
            if customer:
                q = q.filter(Order.customer_name.contains(customer))
            for o in q.all():
                order_ids.add(o.id)

        # 該当する全注文を取得（全商品表示）
        if order_ids:
            orders = Order.query.filter(Order.id.in_(order_ids)).order_by(Order.ordered_at.desc()).all()
            for o in orders:
                results.append(_build_result(o))

    return render_template(
        'order_search.html',
        results=results,
        searched=searched,
        order_no=order_no,
        customer=customer,
        product_code=product_code,
    )
