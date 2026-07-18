from models import db, Inventory


def test_can_create_inventory_row(app):
    inv = Inventory(product_code='TEST-001', inventory_type='即納',
                    quantity=5, reserved_qty=0, available_qty=5)
    db.session.add(inv)
    db.session.commit()

    got = Inventory.query.filter_by(product_code='TEST-001').first()
    assert got is not None
    assert got.quantity == 5
