"""サンプルデータ投入スクリプト"""
from datetime import datetime, date, timedelta
from app import create_app
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.ems import Ems
from models.ems_item import EmsItem
from models.inventory import Inventory
from models.allocation import Allocation
from models.alert import Alert
from models.japan_inventory import JapanInventoryStaging


def seed():
    app = create_app()
    with app.app_context():
        # 既存データをクリア
        db.drop_all()
        db.create_all()

        # === 受注データ ===
        o1 = Order(yahoo_order_id='10114236', ordered_at=datetime(2026, 4, 1, 14, 30),
                    customer_name='奥本 芹奈', status='pending')
        o2 = Order(yahoo_order_id='10114235', ordered_at=datetime(2026, 4, 2, 10, 15),
                    customer_name='田中 美咲', status='pending')
        o3 = Order(yahoo_order_id='10114234', ordered_at=datetime(2026, 4, 3, 9, 0),
                    customer_name='鈴木 太郎', status='pending')
        o4 = Order(yahoo_order_id='10114231', ordered_at=datetime(2026, 4, 5, 11, 20),
                    customer_name='佐藤 花子', status='pending')
        o5 = Order(yahoo_order_id='10114230', ordered_at=datetime(2026, 4, 6, 16, 45),
                    customer_name='奥本 芹奈', status='pending')
        o6 = Order(yahoo_order_id='10114229', ordered_at=datetime(2026, 4, 7, 10, 1),
                    customer_name='宮本 藍', status='pending')
        o7 = Order(yahoo_order_id='10114228', ordered_at=datetime(2026, 4, 7, 12, 0),
                    customer_name='橋本 里桜', status='pending')
        db.session.add_all([o1, o2, o3, o4, o5, o6, o7])
        db.session.flush()

        # === 受注明細 ===
        # o1: 台湾版まんが（お取り寄せ）+ 台湾版まんが（即納）
        oi1 = OrderItem(order_id=o1.id, product_code='TWF0076-CM-20',
                        product_name='台湾版まんが', quantity=1,
                        inventory_type='お取り寄せ', status='provisional_allocated')
        oi2 = OrderItem(order_id=o1.id, product_code='DANDADAN-TWS',
                        product_name='台湾版まんが', quantity=1,
                        inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        # o2: 韓国語の手芸（即納）
        oi3 = OrderItem(order_id=o2.id, product_code='KUMIHIMO01',
                        product_name='韓国語の手芸', quantity=1,
                        inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        # o3: 韓国語参考書（即納）
        oi4 = OrderItem(order_id=o3.id, product_code='KOREAU-C-03-S',
                        product_name='韓国語参考書', quantity=1,
                        inventory_type='即納', status='fully_allocated', allocated_qty=1)
        # o4: 韓国語こども（お取り寄せ）
        oi5 = OrderItem(order_id=o4.id, product_code='KODOMO03',
                        product_name='韓国語こども', quantity=1,
                        inventory_type='お取り寄せ', status='provisional_allocated')
        # o5: 韓国語ぬりえ（お取り寄せ）
        oi6 = OrderItem(order_id=o5.id, product_code='DNURIE71',
                        product_name='韓国語ぬりえ本', quantity=1,
                        inventory_type='お取り寄せ', status='provisional_allocated')
        # o6: ELLE韓国雑誌（即納）
        oi7 = OrderItem(order_id=o6.id, product_code='ELLE2604A',
                        product_name='【A TYPE】韓国雑誌 ELLE', quantity=1,
                        inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        # o6: ミスターブルーグッズ
        oi8 = OrderItem(order_id=o6.id, product_code='MRBLUE40_3',
                        product_name='ミスターブルーグッズ', quantity=4,
                        inventory_type='お取り寄せ', status='fully_allocated', allocated_qty=4)
        # o7: 韓国語参考書3点
        oi9 = OrderItem(order_id=o7.id, product_code='KOREAU-03A',
                        product_name='高麗大おもしろい韓国語2 読む', quantity=1,
                        inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        oi10 = OrderItem(order_id=o7.id, product_code='KOREAU-03B',
                         product_name='高麗大おもしろい韓国語1 書く', quantity=1,
                         inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        oi11 = OrderItem(order_id=o7.id, product_code='KOREAU-03C',
                         product_name='高麗大おもしろい韓国語2 書く', quantity=1,
                         inventory_type='即納', status='allocated_sokunou', allocated_qty=1)
        # o6: ミスターブルーグッズ追加（未発注分）
        oi12 = OrderItem(order_id=o6.id, product_code='MRBLUE40_7',
                         product_name='ミスターブルーグッズ', quantity=30,
                         inventory_type='お取り寄せ', status='pending')

        db.session.add_all([oi1, oi2, oi3, oi4, oi5, oi6, oi7, oi8, oi9, oi10, oi11, oi12])
        db.session.flush()

        # === 発注データ ===
        p1 = Purchase(order_item_id=oi1.id, product_code='TWF0076-CM-20',
                      product_name='台湾版まんが', quantity=1,
                      shop_name='yes24', ordered_at=date(2026, 4, 2), status='ordered')
        p2 = Purchase(order_item_id=oi5.id, product_code='KODOMO03',
                      product_name='韓国語こども', quantity=1,
                      shop_name='알라딘', ordered_at=date(2026, 4, 7), status='arrived')
        p3 = Purchase(order_item_id=oi6.id, product_code='DNURIE71',
                      product_name='韓国語ぬりえ本', quantity=1,
                      shop_name='알라딘', ordered_at=date(2026, 4, 7), status='arrived')
        p4 = Purchase(order_item_id=oi8.id, product_code='MRBLUE40_3',
                      product_name='ミスターブルーグッズ', quantity=4,
                      shop_name='교보문고', ordered_at=date(2026, 4, 7), status='arrived')

        db.session.add_all([p1, p2, p3, p4])
        db.session.flush()

        # === EMS便 ===（到着予定は発送日+3日 = 中2日）
        e1 = Ems(ems_number='EMS240408', shipped_at=date(2026, 4, 8),
                 estimated_arrival=date(2026, 4, 11), arrived_at=date(2026, 4, 11),
                 status='arrived')
        e2 = Ems(ems_number='EMS240412', shipped_at=date(2026, 4, 12),
                 estimated_arrival=date(2026, 4, 15), arrived_at=None,
                 status='in_transit')
        # 遅延デモ用：発送から11日経過・未入荷
        e3 = Ems(ems_number='EMS240403', shipped_at=date(2026, 4, 3),
                 estimated_arrival=date(2026, 4, 6), arrived_at=None,
                 status='in_transit')
        db.session.add_all([e1, e2, e3])
        db.session.flush()

        # === EMS明細 ===
        ei1 = EmsItem(ems_id=e1.id, order_item_id=oi5.id,
                      product_code='KODOMO03', quantity=1)
        ei2 = EmsItem(ems_id=e1.id, order_item_id=oi6.id,
                      product_code='DNURIE71', quantity=1)
        ei3 = EmsItem(ems_id=e1.id, order_item_id=oi8.id,
                      product_code='MRBLUE40_3', quantity=4)
        ei4 = EmsItem(ems_id=e2.id, order_item_id=oi12.id,
                      product_code='MRBLUE40_8', quantity=30)
        ei5 = EmsItem(ems_id=e3.id, order_item_id=oi1.id,
                      product_code='TWF0076-CM-20', quantity=1)
        db.session.add_all([ei1, ei2, ei3, ei4, ei5])
        db.session.flush()

        # === 在庫 ===
        inv1 = Inventory(product_code='DANDADAN-TWS', inventory_type='即納',
                         quantity=3, reserved_qty=1, available_qty=2)
        inv2 = Inventory(product_code='KUMIHIMO01', inventory_type='即納',
                         quantity=2, reserved_qty=1, available_qty=1)
        inv3 = Inventory(product_code='KOREAU-C-03-S', inventory_type='即納',
                         quantity=1, reserved_qty=1, available_qty=0)
        inv4 = Inventory(product_code='ELLE2604A', inventory_type='即納',
                         quantity=5, reserved_qty=1, available_qty=4)
        inv5 = Inventory(product_code='KOREAU-03A', inventory_type='即納',
                         quantity=2, reserved_qty=1, available_qty=1)
        inv6 = Inventory(product_code='KOREAU-03B', inventory_type='即納',
                         quantity=2, reserved_qty=1, available_qty=1)
        inv7 = Inventory(product_code='KOREAU-03C', inventory_type='即納',
                         quantity=2, reserved_qty=1, available_qty=1)
        db.session.add_all([inv1, inv2, inv3, inv4, inv5, inv6, inv7])

        # === 日本在庫仕分け ===
        jis1 = JapanInventoryStaging(ems_item_id=ei1.id, product_code='KODOMO03',
                                     quantity=1, status='waiting')
        jis2 = JapanInventoryStaging(ems_item_id=ei2.id, product_code='DNURIE71',
                                     quantity=1, status='to_japan_stock')
        jis3 = JapanInventoryStaging(ems_item_id=ei3.id, product_code='MRBLUE40_3',
                                     quantity=4, status='assigned_to_order',
                                     assigned_order_item_id=oi8.id)
        db.session.add_all([jis1, jis2, jis3])

        # === アラート ===
        a1 = Alert(alert_type='delay_warning', order_id=o1.id, order_item_id=oi1.id,
                   product_code='TWF0076-CM-20',
                   message='受注 #10114236 — 12日経過・お客様未連絡')
        a2 = Alert(alert_type='korea_ship_missing', order_item_id=oi1.id,
                   product_code='TWF0076-CM-20',
                   message='TWF0076-CM-20 — お取り寄せだがEMS未搭載')
        a3 = Alert(alert_type='japan_ship_missing', product_code='EMS240408',
                   message='EMS #EMS240408 — 4/10入荷・未発送商品あり')
        a4 = Alert(alert_type='purchase_missing', order_item_id=oi12.id,
                   product_code='MRBLUE40_7',
                   message='MRBLUE40_7 (30個) — 受注あるが未発注')
        db.session.add_all([a1, a2, a3, a4])

        db.session.commit()
        print('サンプルデータの投入が完了しました。')
        print(f'  受注: {Order.query.count()}件')
        print(f'  受注明細: {OrderItem.query.count()}件')
        print(f'  発注: {Purchase.query.count()}件')
        print(f'  EMS便: {Ems.query.count()}件')
        print(f'  EMS明細: {EmsItem.query.count()}件')
        print(f'  在庫: {Inventory.query.count()}件')
        print(f'  日本在庫仕分け: {JapanInventoryStaging.query.count()}件')
        print(f'  アラート: {Alert.query.count()}件')


if __name__ == '__main__':
    seed()
