import csv
import io
from flask import Blueprint, render_template, Response
from models import db
from models.ems import Ems
from models.ems_item import EmsItem
from models.order import Order
from models.order_item import OrderItem
from datetime import datetime

bp = Blueprint('ems', __name__, url_prefix='/ems')


@bp.route('/')
def index():
    ems_list = Ems.query.order_by(Ems.shipped_at.desc()).all()
    ems_data = []

    for e in ems_list:
        items = db.session.query(EmsItem, OrderItem, Order).join(
            OrderItem, EmsItem.order_item_id == OrderItem.id
        ).join(
            Order, OrderItem.order_id == Order.id
        ).filter(EmsItem.ems_id == e.id).all()

        for ei, oi, o in items:
            ems_data.append({
                'ems': e,
                'ems_item': ei,
                'order_item': oi,
                'order': o,
            })

    return render_template('ems.html', ems_data=ems_data)


@bp.route('/csv')
def export_csv():
    ems_list = Ems.query.order_by(Ems.shipped_at.desc()).all()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['EMS番号', '発送日', '到着予定日', '入荷日', '商品コード', '数量',
                     '引当先受注番号', 'ステータス'])

    for e in ems_list:
        items = db.session.query(EmsItem, OrderItem, Order).join(
            OrderItem, EmsItem.order_item_id == OrderItem.id
        ).join(
            Order, OrderItem.order_id == Order.id
        ).filter(EmsItem.ems_id == e.id).all()

        for ei, oi, o in items:
            writer.writerow([
                e.ems_number,
                e.shipped_at.strftime('%Y/%m/%d') if e.shipped_at else '',
                e.estimated_arrival.strftime('%Y/%m/%d') if e.estimated_arrival else '',
                e.arrived_at.strftime('%Y/%m/%d') if e.arrived_at else '',
                ei.product_code,
                ei.quantity,
                o.yahoo_order_id,
                e.status_label,
            ])

    output.seek(0)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=ems_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
