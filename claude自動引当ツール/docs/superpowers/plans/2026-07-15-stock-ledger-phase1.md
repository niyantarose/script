# 在庫足し引き台帳 Phase 1 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** すべての在庫増減を追記型台帳 `stock_transactions` に記録し、Yahoo注文の取込で自動マイナス・キャンセル検知で自動プラスされる「足し引きが正確な」日本在庫管理の核を作る。

**Architecture:** 追記型台帳(source_key UNIQUE制約で二重計上を構造的に防止)を唯一の正とし、既存 `Inventory`(inventory_type='即納')の quantity は台帳から再計算されるキャッシュに再定義する。既存の注文取込(`import_yahoo_orders`)と日本到着反映(`/japan/reflect`)の後段に台帳記録をフックする。

**Tech Stack:** Python 3.12 / Flask / Flask-SQLAlchemy (開発=SQLite, 本番=config.py で PostgreSQL/MySQL 切替可) / pytest(新規導入) / Jinja2テンプレート(base.html継承)

**スペック:** `docs/superpowers/specs/2026-07-15-inventory-ledger-pivot-design.md`

## Global Constraints

- 管理対象は日本国内在庫のみ(Inventory の inventory_type='即納' 行が対応)。海外・EMS輸送中は対象外。
- 台帳への訂正は行の書き換え禁止。必ず逆方向の adjust 行を追加する。
- source_key は UNIQUE NOT NULL。自動処理の足し引きは必ず決定的な source_key を持つ(再実行しても1回だけ記録される)。
- 出荷済み注文のキャンセルは自動で在庫を戻さない(アラートのみ)。
- DB列型は SQLite/PostgreSQL/MySQL いずれでも動く型のみ使用(db.String に長さ指定必須、db.Text、db.Integer、db.DateTime、db.Boolean)。
- UIは日本語。テンプレートは `{% extends 'base.html' %}` + 既存CSSクラス(topbar, metric, filter-bar, btn 等)を踏襲。
- コマンドは Windows (PowerShell) 前提。テスト実行は `python -m pytest`。
- 各タスクの最後に必ずコミット。コミットメッセージ末尾に `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>` を付ける。

---

### Task 1: pytest テスト基盤

**Files:**
- Create: `tests/__init__.py` (空ファイル)
- Create: `tests/conftest.py`
- Create: `tests/test_smoke.py`

**Interfaces:**
- Produces: pytest fixture `app`(app_context内でin-memory SQLite + 全テーブル作成済み)、`client`(Flask test client)。以降の全テストがこれを使う。

- [ ] **Step 1: pytest をインストール**

Run: `python -m pip install pytest`
Expected: `Successfully installed pytest-8.x.x` (既にあれば Requirement already satisfied)

- [ ] **Step 2: conftest.py を作成**

`tests/__init__.py` は空ファイルとして作成。`tests/conftest.py`:

```python
import os
import sys

import pytest
from flask import Flask

# プロジェクトルートを import パスに追加
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from models import db as _db  # noqa: E402


@pytest.fixture()
def app():
    """in-memory SQLite で全テーブルを作った素の Flask アプリ。

    app.py の create_app() は .env や APScheduler に依存するため使わず、
    モデル層のテストに必要な最小構成だけ組む。
    """
    flask_app = Flask(
        __name__,
        template_folder=os.path.join(ROOT, 'templates'),
        static_folder=os.path.join(ROOT, 'static'),
    )
    flask_app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
    flask_app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    flask_app.config['TESTING'] = True
    _db.init_app(flask_app)
    with flask_app.app_context():
        _db.create_all()
        yield flask_app
        _db.session.rollback()
        _db.drop_all()


@pytest.fixture()
def client(app):
    return app.test_client()
```

- [ ] **Step 3: スモークテストを書く**

`tests/test_smoke.py`:

```python
from models import db, Inventory


def test_can_create_inventory_row(app):
    inv = Inventory(product_code='TEST-001', inventory_type='即納',
                    quantity=5, reserved_qty=0, available_qty=5)
    db.session.add(inv)
    db.session.commit()

    got = Inventory.query.filter_by(product_code='TEST-001').first()
    assert got is not None
    assert got.quantity == 5
```

- [ ] **Step 4: テスト実行**

Run: `python -m pytest tests/test_smoke.py -v`
Expected: `1 passed`

- [ ] **Step 5: コミット**

```powershell
git add tests/__init__.py tests/conftest.py tests/test_smoke.py
git commit -m "test(zaiko): pytestテスト基盤を追加（in-memory SQLite fixture）"
```

---

### Task 2: StockTransaction / MallSku モデル

**Files:**
- Create: `models/stock_transaction.py`
- Create: `models/mall_sku.py`
- Modify: `models/__init__.py` (import追加)
- Test: `tests/test_models_ledger.py`

**Interfaces:**
- Produces:
  - `StockTransaction(product_code, product_sub_code, tx_type, qty, ref_type, ref_id, source_key, reason, created_at)` — `source_key` UNIQUE NOT NULL
  - `StockTransaction.TX_TYPE_LABELS: dict[str, str]` — キー: `receive / order_out / cancel_return / manual_in / manual_out / adjust`
  - `MallSku(mall, external_code, external_sub_code, product_code, product_sub_code)` — UNIQUE(mall, external_code)

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_models_ledger.py`:

```python
import pytest
from sqlalchemy.exc import IntegrityError

from models import db, StockTransaction, MallSku


def test_stock_transaction_insert(app):
    tx = StockTransaction(product_code='P-001', tx_type='manual_in',
                          qty=3, source_key='manual:abc', reason='テスト入庫')
    db.session.add(tx)
    db.session.commit()
    assert StockTransaction.query.count() == 1


def test_source_key_unique_blocks_duplicate(app):
    db.session.add(StockTransaction(product_code='P-001', tx_type='order_out',
                                    qty=-1, source_key='yahoo:o1:1:out'))
    db.session.commit()
    db.session.add(StockTransaction(product_code='P-001', tx_type='order_out',
                                    qty=-1, source_key='yahoo:o1:1:out'))
    with pytest.raises(IntegrityError):
        db.session.commit()
    db.session.rollback()
    assert StockTransaction.query.count() == 1


def test_mall_sku_unique_per_mall(app):
    db.session.add(MallSku(mall='yahoo', external_code='EXT-1', product_code='P-001'))
    db.session.commit()
    db.session.add(MallSku(mall='yahoo', external_code='EXT-1', product_code='P-002'))
    with pytest.raises(IntegrityError):
        db.session.commit()
    db.session.rollback()
    # 別モールなら同じ external_code でも登録できる
    db.session.add(MallSku(mall='amazon', external_code='EXT-1', product_code='P-001'))
    db.session.commit()
    assert MallSku.query.count() == 2
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_models_ledger.py -v`
Expected: FAIL — `ImportError: cannot import name 'StockTransaction'`

- [ ] **Step 3: モデルを実装**

`models/stock_transaction.py`:

```python
from datetime import datetime
from models import db


class StockTransaction(db.Model):
    """在庫トランザクション台帳（追記型）。

    現在庫 = この台帳の qty 合計。行の書き換え・削除は禁止で、
    訂正は逆方向の adjust 行を追加する。
    """
    __tablename__ = 'stock_transactions'

    id               = db.Column(db.Integer, primary_key=True)
    product_code     = db.Column(db.String(100), nullable=False, index=True)
    product_sub_code = db.Column(db.String(100), nullable=True)
    tx_type          = db.Column(db.String(30), nullable=False)
    qty              = db.Column(db.Integer, nullable=False)  # 符号付き（＋入庫/−出庫）
    ref_type         = db.Column(db.String(30), nullable=True)   # order_item / japan_staging / stocktake / manual
    ref_id           = db.Column(db.Integer, nullable=True)
    # 発生元の一意キー。UNIQUE制約が二重計上を構造的に防ぐ
    source_key       = db.Column(db.String(200), nullable=False, unique=True)
    reason           = db.Column(db.Text, nullable=True)
    created_at       = db.Column(db.DateTime, default=datetime.now)

    TX_TYPE_LABELS = {
        'receive':       '入庫（仕入れ到着）',
        'order_out':     '注文出庫',
        'cancel_return': 'キャンセル戻し',
        'manual_in':     '手動入庫',
        'manual_out':    '手動出庫',
        'adjust':        '調整（棚卸・訂正）',
    }

    @property
    def tx_type_label(self):
        return self.TX_TYPE_LABELS.get(self.tx_type, self.tx_type)

    def to_dict(self):
        return {
            'id': self.id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'tx_type': self.tx_type,
            'tx_type_label': self.tx_type_label,
            'qty': self.qty,
            'ref_type': self.ref_type or '',
            'ref_id': self.ref_id,
            'source_key': self.source_key,
            'reason': self.reason or '',
            'created_at': self.created_at.strftime('%Y/%m/%d %H:%M') if self.created_at else '',
        }
```

`models/mall_sku.py`:

```python
from datetime import datetime
from models import db


class MallSku(db.Model):
    """モール側商品コード → 内部商品コードのマッピング（Phase 4 の土台）。"""
    __tablename__ = 'mall_skus'
    __table_args__ = (
        db.UniqueConstraint('mall', 'external_code', name='uq_mall_external'),
    )

    id                = db.Column(db.Integer, primary_key=True)
    mall              = db.Column(db.String(20), nullable=False)  # yahoo/amazon/qoo10/mercari/tiktok
    external_code     = db.Column(db.String(100), nullable=False)
    external_sub_code = db.Column(db.String(100), nullable=True)
    product_code      = db.Column(db.String(100), nullable=False, index=True)
    product_sub_code  = db.Column(db.String(100), nullable=True)
    created_at        = db.Column(db.DateTime, default=datetime.now)
```

`models/__init__.py` の末尾に追加:

```python
from models.stock_transaction import StockTransaction
from models.mall_sku import MallSku
```

- [ ] **Step 4: テストが通ることを確認**

Run: `python -m pytest tests/test_models_ledger.py -v`
Expected: `3 passed`

- [ ] **Step 5: コミット**

```powershell
git add models/stock_transaction.py models/mall_sku.py models/__init__.py tests/test_models_ledger.py
git commit -m "feat(zaiko): 在庫トランザクション台帳とモールSKUマッピングのモデルを追加"
```

---

### Task 3: 台帳サービスの核（record_transaction / get_balance / recalc_inventory）

**Files:**
- Create: `services/stock_ledger.py`
- Test: `tests/test_stock_ledger.py`

**Interfaces:**
- Produces:
  - `record_transaction(product_code, tx_type, qty, source_key, *, product_sub_code=None, ref_type=None, ref_id=None, reason=None) -> tuple[StockTransaction | None, bool]` — 戻り値は `(tx, created)`。source_key 既存なら `(既存tx, False)` で何もしない。qty の符号が tx_type と矛盾したら `ValueError`。記録後に自動で `recalc_inventory` を呼ぶ。**commit はしない**(呼び出し側の責務)。
  - `get_balance(product_code) -> int` — 台帳合計。
  - `recalc_inventory(product_code, product_sub_code=None) -> Inventory` — 即納 Inventory 行の quantity を台帳合計で上書きし、available_qty = max(0, quantity - reserved_qty) に更新。行がなければ作る。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_stock_ledger.py`:

```python
import pytest

from models import db, Inventory, StockTransaction
from services.stock_ledger import record_transaction, get_balance, recalc_inventory


def test_record_and_balance(app):
    tx, created = record_transaction('P-001', 'manual_in', 5, 'manual:t1', reason='入庫')
    db.session.commit()
    assert created is True
    assert get_balance('P-001') == 5

    record_transaction('P-001', 'manual_out', -2, 'manual:t2', reason='破損')
    db.session.commit()
    assert get_balance('P-001') == 3


def test_same_source_key_records_once(app):
    record_transaction('P-001', 'order_out', -1, 'yahoo:o1:1:out')
    db.session.commit()
    tx, created = record_transaction('P-001', 'order_out', -1, 'yahoo:o1:1:out')
    db.session.commit()
    assert created is False
    assert StockTransaction.query.count() == 1
    assert get_balance('P-001') == -1


def test_sign_validation(app):
    with pytest.raises(ValueError):
        record_transaction('P-001', 'manual_in', -3, 'manual:bad1')  # 入庫は＋のみ
    with pytest.raises(ValueError):
        record_transaction('P-001', 'order_out', 1, 'manual:bad2')   # 出庫は−のみ
    with pytest.raises(ValueError):
        record_transaction('P-001', 'adjust', 0, 'manual:bad3')      # 0は禁止


def test_recalc_updates_sokunou_inventory(app):
    inv = Inventory(product_code='P-002', inventory_type='即納',
                    quantity=10, reserved_qty=4, available_qty=6)
    db.session.add(inv)
    db.session.commit()

    record_transaction('P-002', 'adjust', 10, 'seed:P-002', reason='期首残高')
    record_transaction('P-002', 'order_out', -3, 'yahoo:o2:9:out')
    db.session.commit()

    got = Inventory.query.filter_by(product_code='P-002', inventory_type='即納').first()
    assert got.quantity == 7            # 10 - 3
    assert got.available_qty == 3       # 7 - reserved 4


def test_recalc_creates_missing_inventory_row(app):
    record_transaction('P-NEW', 'manual_in', 2, 'manual:t3')
    db.session.commit()
    got = Inventory.query.filter_by(product_code='P-NEW', inventory_type='即納').first()
    assert got is not None
    assert got.quantity == 2
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_stock_ledger.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'services.stock_ledger'`

- [ ] **Step 3: サービスを実装**

`services/stock_ledger.py`:

```python
"""在庫足し引き台帳サービス。

すべての在庫増減はここを通して stock_transactions に記録する。
source_key の UNIQUE 制約により、同じ発生源の足し引きは何度呼んでも1回だけ。
record_transaction / 各同期関数は commit しない（呼び出し側で commit する）。
"""
from sqlalchemy import func

from models import db, Inventory, StockTransaction

# tx_type ごとの符号ルール
_POSITIVE_TYPES = {'receive', 'cancel_return', 'manual_in'}
_NEGATIVE_TYPES = {'order_out', 'manual_out'}
_ANY_SIGN_TYPES = {'adjust'}


def record_transaction(product_code, tx_type, qty, source_key, *,
                       product_sub_code=None, ref_type=None, ref_id=None,
                       reason=None):
    """台帳に1行記録し、即納Inventoryキャッシュを再計算する。

    Returns:
        (tx, created): source_key が既存なら (既存tx, False)
    Raises:
        ValueError: tx_type と qty の符号が矛盾、または qty=0
    """
    if tx_type in _POSITIVE_TYPES and qty <= 0:
        raise ValueError(f'{tx_type} の qty は正の値のみ: {qty}')
    if tx_type in _NEGATIVE_TYPES and qty >= 0:
        raise ValueError(f'{tx_type} の qty は負の値のみ: {qty}')
    if tx_type in _ANY_SIGN_TYPES and qty == 0:
        raise ValueError('adjust の qty に 0 は指定できません')
    if tx_type not in (_POSITIVE_TYPES | _NEGATIVE_TYPES | _ANY_SIGN_TYPES):
        raise ValueError(f'不明な tx_type: {tx_type}')

    existing = StockTransaction.query.filter_by(source_key=source_key).first()
    if existing:
        return existing, False

    tx = StockTransaction(
        product_code=product_code,
        product_sub_code=product_sub_code,
        tx_type=tx_type,
        qty=qty,
        ref_type=ref_type,
        ref_id=ref_id,
        source_key=source_key,
        reason=reason,
    )
    db.session.add(tx)
    db.session.flush()
    recalc_inventory(product_code, product_sub_code)
    return tx, True


def get_balance(product_code):
    """台帳上の現在庫（qty 合計）。"""
    total = db.session.query(func.coalesce(func.sum(StockTransaction.qty), 0)) \
        .filter(StockTransaction.product_code == product_code).scalar()
    return int(total or 0)


def recalc_inventory(product_code, product_sub_code=None):
    """即納 Inventory 行の quantity を台帳合計で上書きする（キャッシュ更新）。"""
    balance = get_balance(product_code)
    inv = Inventory.query.filter_by(
        product_code=product_code, inventory_type='即納').first()
    if not inv:
        inv = Inventory(
            product_code=product_code,
            product_sub_code=product_sub_code,
            inventory_type='即納',
            quantity=0, reserved_qty=0, available_qty=0,
        )
        db.session.add(inv)
    inv.quantity = balance
    inv.available_qty = max(0, balance - (inv.reserved_qty or 0))
    return inv
```

- [ ] **Step 4: テストが通ることを確認**

Run: `python -m pytest tests/test_stock_ledger.py -v`
Expected: `5 passed`

- [ ] **Step 5: 全テスト実行（回帰確認）**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 6: コミット**

```powershell
git add services/stock_ledger.py tests/test_stock_ledger.py
git commit -m "feat(zaiko): 台帳サービスの核を追加（記録・残高・即納キャッシュ再計算）"
```

---

### Task 4: 期首残高シードスクリプト

**Files:**
- Create: `scripts/seed_ledger_opening.py`
- Test: `tests/test_seed_opening.py`

**Interfaces:**
- Consumes: `record_transaction` (Task 3)
- Produces:
  - `seed_opening_balances() -> dict` — `{'seeded': n, 'skipped': n}`。即納 Inventory 全行に `adjust` / source_key `seed:{product_code}` で期首残高を記録。quantity=0 の行はスキップ。
  - `seed_yahoo_mall_skus() -> dict` — `{'created': n, 'skipped': n}`。inventory_type='yahoo' の全行から MallSku(mall='yahoo', external=内部コード1:1) を作成。
  - CLI: `python scripts/seed_ledger_opening.py` で両方実行(何度実行しても安全)。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_seed_opening.py`:

```python
from models import db, Inventory, StockTransaction, MallSku
from scripts.seed_ledger_opening import seed_opening_balances, seed_yahoo_mall_skus
from services.stock_ledger import get_balance


def _setup_rows():
    db.session.add(Inventory(product_code='P-A', inventory_type='即納',
                             quantity=7, reserved_qty=0, available_qty=7))
    db.session.add(Inventory(product_code='P-B', inventory_type='即納',
                             quantity=0, reserved_qty=0, available_qty=0))
    db.session.add(Inventory(product_code='P-A', inventory_type='yahoo',
                             quantity=7, yahoo_stock=7))
    db.session.commit()


def test_seed_opening_is_idempotent(app):
    _setup_rows()
    r1 = seed_opening_balances()
    db.session.commit()
    assert r1['seeded'] == 1          # P-A のみ（P-B は 0 なのでスキップ）
    assert get_balance('P-A') == 7

    r2 = seed_opening_balances()      # 2回目は何も増えない
    db.session.commit()
    assert r2['seeded'] == 0
    assert get_balance('P-A') == 7
    assert StockTransaction.query.count() == 1


def test_seed_yahoo_mall_skus(app):
    _setup_rows()
    r1 = seed_yahoo_mall_skus()
    db.session.commit()
    assert r1['created'] == 1
    ms = MallSku.query.filter_by(mall='yahoo', external_code='P-A').first()
    assert ms.product_code == 'P-A'

    r2 = seed_yahoo_mall_skus()
    db.session.commit()
    assert r2['created'] == 0
    assert MallSku.query.count() == 1
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_seed_opening.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.seed_ledger_opening'`

- [ ] **Step 3: スクリプトを実装**

`scripts/__init__.py` が無ければ空ファイルで作成。`scripts/seed_ledger_opening.py`:

```python
"""台帳の期首残高と Yahoo モールSKUマッピングをシードする。

使い方（プロジェクトルートで）:
    python scripts/seed_ledger_opening.py

何度実行しても安全（source_key / UNIQUE制約で冪等）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models import db, Inventory, MallSku
from services.stock_ledger import record_transaction


def seed_opening_balances():
    """即納 Inventory の現在値を期首残高として台帳に記録する。"""
    seeded = skipped = 0
    rows = Inventory.query.filter_by(inventory_type='即納').all()
    for inv in rows:
        qty = inv.quantity or 0
        if qty == 0:
            skipped += 1
            continue
        _, created = record_transaction(
            inv.product_code, 'adjust', qty, f'seed:{inv.product_code}',
            product_sub_code=inv.product_sub_code,
            ref_type='manual', reason='期首残高（台帳導入時の初期値）',
        )
        seeded += 1 if created else 0
        skipped += 0 if created else 1
    return {'seeded': seeded, 'skipped': skipped}


def seed_yahoo_mall_skus():
    """Yahoo 在庫行から MallSku を1:1で作成する。"""
    created = skipped = 0
    rows = Inventory.query.filter_by(inventory_type='yahoo').all()
    for inv in rows:
        exists = MallSku.query.filter_by(
            mall='yahoo', external_code=inv.product_code).first()
        if exists:
            skipped += 1
            continue
        db.session.add(MallSku(
            mall='yahoo',
            external_code=inv.product_code,
            external_sub_code=inv.product_sub_code,
            product_code=inv.product_code,
            product_sub_code=inv.product_sub_code,
        ))
        created += 1
    return {'created': created, 'skipped': skipped}


if __name__ == '__main__':
    from app import app
    with app.app_context():
        r1 = seed_opening_balances()
        r2 = seed_yahoo_mall_skus()
        db.session.commit()
        print(f'期首残高: {r1} / YahooモールSKU: {r2}')
```

- [ ] **Step 4: テストが通ることを確認**

Run: `python -m pytest tests/test_seed_opening.py -v`
Expected: `2 passed`

- [ ] **Step 5: コミット**

```powershell
git add scripts/__init__.py scripts/seed_ledger_opening.py tests/test_seed_opening.py
git commit -m "feat(zaiko): 期首残高とYahooモールSKUのシードスクリプトを追加"
```

**注意:** 本番DBへの実行は Phase 1 完了後の切替作業で行う(このタスクではテストのみ)。

---

### Task 5: 注文→自動マイナス、キャンセル→自動プラス

**Files:**
- Modify: `services/stock_ledger.py` (関数追加)
- Modify: `routes/import_data.py:52-61` (取込後のフック追加)
- Modify: `models/alert.py:21-28` (ALERT_TYPE_LABELS に2種追加)
- Test: `tests/test_order_ledger_sync.py`

**Interfaces:**
- Consumes: `record_transaction`, `get_balance` (Task 3)
- Produces:
  - `apply_order_out(item_ids: list[int]) -> int` — 新規 OrderItem のみ対象。キャンセル済み(yahoo_order_status=='4')・出荷済み(yahoo_ship_status in ('2','3'))の注文はスキップ。source_key: `yahoo:{yahoo_order_id}:{order_item.id}:out`。戻り値=記録件数。
  - `sync_cancel_returns() -> dict` — `{'returned': n, 'alerted': n}`。キャンセル注文のうち out 記録済みの明細に cancel_return(source_key: `yahoo:{yahoo_order_id}:{order_item.id}:return`)を記録。出荷済みキャンセルは戻さず Alert(alert_type='shipped_cancel')を1件だけ作成。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_order_ledger_sync.py`:

```python
from datetime import datetime

from models import db, Order, OrderItem, Alert, StockTransaction
from services.stock_ledger import (
    apply_order_out, sync_cancel_returns, get_balance, record_transaction,
)


def _make_order(yahoo_order_id, order_status='2', ship_status='0'):
    o = Order(yahoo_order_id=yahoo_order_id, ordered_at=datetime.now(),
              yahoo_order_status=order_status, yahoo_ship_status=ship_status)
    db.session.add(o)
    db.session.flush()
    return o


def _make_item(order, product_code='P-001', qty=2):
    oi = OrderItem(order_id=order.id, product_code=product_code,
                   quantity=qty, inventory_type='pending', status='pending')
    db.session.add(oi)
    db.session.flush()
    return oi


def test_order_out_subtracts_once(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-1')
    oi = _make_item(o, qty=2)
    db.session.commit()

    assert apply_order_out([oi.id]) == 1
    db.session.commit()
    assert get_balance('P-001') == 8

    # 再実行しても二重に引かれない
    assert apply_order_out([oi.id]) == 0
    db.session.commit()
    assert get_balance('P-001') == 8


def test_cancelled_new_order_is_skipped(app):
    o = _make_order('order-2', order_status='4')  # 取込時点で既にキャンセル
    oi = _make_item(o)
    db.session.commit()
    assert apply_order_out([oi.id]) == 0
    assert get_balance('P-001') == 0


def test_cancel_return_restores_once(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-3')
    oi = _make_item(o, qty=3)
    db.session.commit()
    apply_order_out([oi.id])
    db.session.commit()
    assert get_balance('P-001') == 7

    o.yahoo_order_status = '4'  # キャンセルに変化
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r['returned'] == 1
    assert get_balance('P-001') == 10

    r2 = sync_cancel_returns()  # 再実行しても戻しは1回だけ
    db.session.commit()
    assert r2['returned'] == 0
    assert get_balance('P-001') == 10


def test_shipped_cancel_alerts_instead_of_return(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-4')
    oi = _make_item(o, qty=1)
    db.session.commit()
    apply_order_out([oi.id])
    db.session.commit()

    o.yahoo_order_status = '4'
    o.yahoo_ship_status = '3'  # 出荷済み
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r['returned'] == 0
    assert r['alerted'] == 1
    assert get_balance('P-001') == 9  # 戻っていない
    assert Alert.query.filter_by(alert_type='shipped_cancel').count() == 1

    sync_cancel_returns()  # アラートも重複しない
    db.session.commit()
    assert Alert.query.filter_by(alert_type='shipped_cancel').count() == 1


def test_cancel_without_out_does_nothing(app):
    # 台帳導入前の古いキャンセル注文（out記録なし）には何もしない
    o = _make_order('order-5', order_status='4')
    _make_item(o)
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r == {'returned': 0, 'alerted': 0}
    assert StockTransaction.query.count() == 0
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_order_ledger_sync.py -v`
Expected: FAIL — `ImportError: cannot import name 'apply_order_out'`

- [ ] **Step 3: サービス関数を実装**

`services/stock_ledger.py` の import 部を差し替え、末尾に追加:

```python
from models import db, Alert, Inventory, Order, OrderItem, StockTransaction
```

```python
def apply_order_out(item_ids):
    """新規取込された注文明細の出庫を台帳に記録する。

    対象は item_ids で渡された明細のみ（過去注文を遡って引かない）。
    キャンセル済み・出荷済みの注文はスキップ。
    """
    recorded = 0
    for item_id in item_ids or []:
        item = OrderItem.query.get(item_id)
        if not item:
            continue
        order = Order.query.get(item.order_id)
        if not order:
            continue
        if order.yahoo_order_status == '4':          # キャンセル済み
            continue
        if order.yahoo_ship_status in ('2', '3'):    # 出荷処理中・出荷済み
            continue
        _, created = record_transaction(
            item.product_code, 'order_out', -item.quantity,
            f'yahoo:{order.yahoo_order_id}:{item.id}:out',
            product_sub_code=item.product_sub_code,
            ref_type='order_item', ref_id=item.id,
        )
        recorded += 1 if created else 0
    return recorded


def sync_cancel_returns():
    """キャンセル注文の在庫戻しを台帳に記録する。

    - 台帳に out 記録がある明細だけが対象（導入前の古いキャンセルは無視）
    - 出荷済みキャンセルは自動で戻さず shipped_cancel アラートを1件作成
    """
    returned = alerted = 0
    cancelled_orders = Order.query.filter_by(yahoo_order_status='4').all()
    for order in cancelled_orders:
        for item in order.items:
            out_key = f'yahoo:{order.yahoo_order_id}:{item.id}:out'
            ret_key = f'yahoo:{order.yahoo_order_id}:{item.id}:return'
            has_out = StockTransaction.query.filter_by(source_key=out_key).first()
            if not has_out:
                continue
            has_return = StockTransaction.query.filter_by(source_key=ret_key).first()
            if has_return:
                continue

            if order.yahoo_ship_status in ('2', '3'):
                # 出荷済みキャンセル: 自動で戻さずアラート（重複作成しない）
                exists = Alert.query.filter_by(
                    alert_type='shipped_cancel', order_item_id=item.id).first()
                if not exists:
                    db.session.add(Alert(
                        alert_type='shipped_cancel',
                        order_id=order.id,
                        order_item_id=item.id,
                        product_code=item.product_code,
                        message=(f'注文 {order.yahoo_order_id} は出荷済みのまま'
                                 f'キャンセルされました。返品到着後に手動入庫してください'
                                 f'（{item.product_code} × {item.quantity}）。'),
                    ))
                    alerted += 1
                continue

            _, created = record_transaction(
                item.product_code, 'cancel_return', item.quantity, ret_key,
                product_sub_code=item.product_sub_code,
                ref_type='order_item', ref_id=item.id,
                reason=f'注文 {order.yahoo_order_id} キャンセル',
            )
            returned += 1 if created else 0
    return {'returned': returned, 'alerted': alerted}
```

- [ ] **Step 4: テストが通ることを確認**

Run: `python -m pytest tests/test_order_ledger_sync.py -v`
Expected: `5 passed`

- [ ] **Step 5: アラート種別ラベルを追加**

`models/alert.py` の `ALERT_TYPE_LABELS` に2行追加:

```python
    ALERT_TYPE_LABELS = {
        'purchase_missing': '発注漏れ',
        'korea_ship_missing': '韓国発送漏れ',
        'japan_arrival_missing': '日本入荷漏れ',
        'japan_ship_missing': '発送漏れ',
        'stock_shortage': '在庫不足',
        'delay_warning': '遅延警告',
        'shipped_cancel': '出荷済みキャンセル（要手動戻し）',
        'ledger_mismatch': '台帳とキャッシュの不一致',
    }
```

- [ ] **Step 6: 注文取込にフックを追加**

`routes/import_data.py` の `import_yahoo_orders` 内、自動引当ブロック(52-61行付近)の直後・`return jsonify` の前に追加:

```python
        # ── 在庫台帳: 注文出庫・キャンセル戻しを記録 ──────────────────
        from services.stock_ledger import apply_order_out, sync_cancel_returns
        ledger_out = apply_order_out(new_item_ids)
        ledger_ret = sync_cancel_returns()
        db.session.commit()
```

`return jsonify({...})` に2キー追加し、message 末尾にも追記:

```python
            'ledger_out':      ledger_out,
            'ledger_returns':  ledger_ret,
```

message の f-string 末尾に追加:

```python
                        f' | 台帳: 出庫{ledger_out}件 / 戻し{ledger_ret["returned"]}件'
```

- [ ] **Step 7: 全テスト実行**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 8: コミット**

```powershell
git add services/stock_ledger.py routes/import_data.py models/alert.py tests/test_order_ledger_sync.py
git commit -m "feat(zaiko): 注文取込で自動出庫・キャンセル検知で自動戻しを台帳記録"
```

---

### Task 6: 日本到着(reflect)を台帳経由の入庫に変更

**Files:**
- Modify: `routes/japan_inventory.py:149-181` (`reflect_to_yahoo`)
- Test: `tests/test_japan_reflect_ledger.py`

**Interfaces:**
- Consumes: `record_transaction` (Task 3)
- Produces: `/japan/reflect` POST が staging ごとに `receive` tx (source_key: `japan_staging:{jis.id}`) を記録し、Inventory 更新は recalc に委ねる。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_japan_reflect_ledger.py`:

```python
import pytest

from models import db, Inventory, JapanInventoryStaging, StockTransaction
from services.stock_ledger import get_balance


@pytest.fixture()
def jclient(app):
    from routes.japan_inventory import bp
    app.register_blueprint(bp)
    return app.test_client()


def _make_staging(product_code='P-001', qty=4):
    # SQLite は FK 未強制なので ems_item_id はダミーIDで良い
    jis = JapanInventoryStaging(ems_item_id=999, product_code=product_code,
                                quantity=qty, status='to_japan_stock')
    db.session.add(jis)
    db.session.commit()
    return jis


def test_reflect_records_receive_tx(app, jclient):
    jis = _make_staging(qty=4)
    res = jclient.post('/japan/reflect')
    assert res.status_code == 200
    assert res.get_json()['reflected_count'] == 1

    assert get_balance('P-001') == 4
    inv = Inventory.query.filter_by(product_code='P-001', inventory_type='即納').first()
    assert inv.quantity == 4

    tx = StockTransaction.query.filter_by(source_key=f'japan_staging:{jis.id}').first()
    assert tx is not None and tx.tx_type == 'receive'


def test_reflect_is_idempotent(app, jclient):
    jis = _make_staging(qty=4)
    jclient.post('/japan/reflect')
    # ステータスを強制的に戻して再実行しても、台帳は二重計上しない
    jis.status = 'to_japan_stock'
    db.session.commit()
    jclient.post('/japan/reflect')
    assert get_balance('P-001') == 4
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_japan_reflect_ledger.py -v`
Expected: FAIL — `assert get_balance('P-001') == 4` が 0 で失敗(台帳未記録のため)

**注意:** `/japan/reflect` の URL prefix が違う場合は `routes/japan_inventory.py` 冒頭の `Blueprint(..., url_prefix=...)` を確認し、テスト内のURLを実際の prefix に合わせること。

- [ ] **Step 3: reflect を台帳経由に書き換え**

`routes/japan_inventory.py` の `reflect_to_yahoo` 内、Inventory を直接加算しているブロック(157-174行付近)を以下に差し替え:

```python
    from services.stock_ledger import record_transaction

    for jis in targets:
        # 台帳に入庫を記録（即納Inventoryのquantityはrecalcが更新する）
        record_transaction(
            jis.product_code, 'receive', jis.quantity,
            f'japan_staging:{jis.id}',
            product_sub_code=jis.product_sub_code,
            ref_type='japan_staging', ref_id=jis.id,
            reason='日本到着反映',
        )

        jis.status = 'reflected'
        jis.reflected_at = datetime.now()
        reflected_count += 1
```

(元の `inv = Inventory.query.filter_by(...)` 〜 `db.session.add(inv)` のブロックは削除する)

- [ ] **Step 4: テストが通ることを確認**

Run: `python -m pytest tests/test_japan_reflect_ledger.py -v`
Expected: `2 passed`

- [ ] **Step 5: 全テスト実行**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 6: コミット**

```powershell
git add routes/japan_inventory.py tests/test_japan_reflect_ledger.py
git commit -m "feat(zaiko): 日本到着反映を台帳経由の入庫記録に変更（二重計上防止）"
```

---

### Task 7: 手動足し引き・履歴のAPI

**Files:**
- Create: `routes/ledger.py`
- Modify: `app.py:22-40` (blueprint登録)
- Test: `tests/test_ledger_routes.py`

**Interfaces:**
- Consumes: `record_transaction`, `get_balance` (Task 3)
- Produces:
  - `GET /ledger/` — 台帳ページ(Task 8 のテンプレートを描画)
  - `GET /ledger/api/balances?q=<検索語>` — `{'items': [{product_code, product_name, ledger_qty, location, reserved_qty, available_qty}]}`(即納 Inventory 行ベース、商品コード/商品名の部分一致)
  - `POST /ledger/api/tx` — body: `{product_code, tx_type(manual_in|manual_out|adjust), qty(正の数で送る), reason}`。manual_out は qty を負に変換して記録。manual_out/adjust は reason 必須(400)。source_key は `manual:{uuid4}`。
  - `GET /ledger/api/history/<product_code>` — `{'items': [tx.to_dict() 新しい順 最大200件]}`

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_ledger_routes.py`:

```python
import pytest

from models import db, Inventory
from services.stock_ledger import get_balance, record_transaction


@pytest.fixture()
def lclient(app):
    from routes.ledger import bp
    app.register_blueprint(bp)
    return app.test_client()


def test_manual_in_and_out(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': 5,
        'reason': '手動入庫テスト'})
    assert res.status_code == 200
    assert get_balance('P-001') == 5

    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_out', 'qty': 2,
        'reason': '破損'})
    assert res.status_code == 200
    assert get_balance('P-001') == 3


def test_manual_out_requires_reason(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_out', 'qty': 1, 'reason': ''})
    assert res.status_code == 400
    assert get_balance('P-001') == 0


def test_invalid_qty_rejected(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': 0})
    assert res.status_code == 400
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': -3})
    assert res.status_code == 400


def test_balances_search(app, lclient):
    db.session.add(Inventory(product_code='ABC-1', product_name='ダンダダン 1巻',
                             inventory_type='即納', quantity=3, reserved_qty=0,
                             available_qty=3, location='A-1'))
    db.session.commit()
    res = lclient.get('/ledger/api/balances?q=ダンダ')
    data = res.get_json()
    assert len(data['items']) == 1
    assert data['items'][0]['product_code'] == 'ABC-1'
    assert data['items'][0]['location'] == 'A-1'


def test_history(app, lclient):
    record_transaction('P-001', 'manual_in', 5, 'manual:h1', reason='入庫1')
    record_transaction('P-001', 'manual_out', -1, 'manual:h2', reason='出庫1')
    db.session.commit()
    res = lclient.get('/ledger/api/history/P-001')
    items = res.get_json()['items']
    assert len(items) == 2
    assert items[0]['reason'] == '出庫1'  # 新しい順
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_ledger_routes.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'routes.ledger'`

- [ ] **Step 3: ルートを実装**

`routes/ledger.py`:

```python
import uuid

from flask import Blueprint, jsonify, render_template, request

from models import db, Inventory
from models.stock_transaction import StockTransaction
from services.stock_ledger import record_transaction

bp = Blueprint('ledger', __name__, url_prefix='/ledger')

_MANUAL_TYPES = {'manual_in', 'manual_out', 'adjust'}
_REASON_REQUIRED = {'manual_out', 'adjust'}


@bp.route('/')
def index():
    return render_template('ledger.html', active_page='ledger')


@bp.route('/api/balances')
def api_balances():
    q = (request.args.get('q') or '').strip()
    query = Inventory.query.filter_by(inventory_type='即納')
    if q:
        like = f'%{q}%'
        query = query.filter(db.or_(
            Inventory.product_code.ilike(like),
            Inventory.product_name.ilike(like),
        ))
    rows = query.order_by(Inventory.product_code).limit(500).all()
    return jsonify({'items': [{
        'product_code': r.product_code,
        'product_name': r.product_name or '',
        'ledger_qty': r.quantity,
        'reserved_qty': r.reserved_qty or 0,
        'available_qty': r.available_qty or 0,
        'location': r.location or '',
    } for r in rows]})


@bp.route('/api/tx', methods=['POST'])
def api_tx():
    body = request.get_json(silent=True) or {}
    product_code = (body.get('product_code') or '').strip()
    tx_type = body.get('tx_type')
    reason = (body.get('reason') or '').strip()
    try:
        qty = int(body.get('qty', 0))
    except (TypeError, ValueError):
        qty = 0

    if not product_code:
        return jsonify({'status': 'error', 'message': '商品コードが必要です'}), 400
    if tx_type not in _MANUAL_TYPES:
        return jsonify({'status': 'error', 'message': f'不正な種別: {tx_type}'}), 400
    if qty <= 0 and tx_type != 'adjust':
        return jsonify({'status': 'error', 'message': '数量は1以上を指定してください'}), 400
    if qty == 0:
        return jsonify({'status': 'error', 'message': '数量に0は指定できません'}), 400
    if tx_type in _REASON_REQUIRED and not reason:
        return jsonify({'status': 'error', 'message': '出庫・調整には理由が必要です'}), 400

    signed_qty = -qty if tx_type == 'manual_out' else qty
    try:
        tx, _ = record_transaction(
            product_code, tx_type, signed_qty, f'manual:{uuid.uuid4()}',
            ref_type='manual', reason=reason or None,
        )
        db.session.commit()
    except ValueError as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 400
    return jsonify({'status': 'ok', 'tx': tx.to_dict()})


@bp.route('/api/history/<path:product_code>')
def api_history(product_code):
    rows = StockTransaction.query.filter_by(product_code=product_code) \
        .order_by(StockTransaction.id.desc()).limit(200).all()
    return jsonify({'items': [r.to_dict() for r in rows]})
```

**注意:** adjust の負数指定はUIから `qty` に正負どちらも送れるようにするため、adjust のときのみ `qty` の符号をそのまま使う(上のコードは qty<=0 を adjust では許容し、qty==0 のみ拒否している)。

- [ ] **Step 4: app.py に blueprint を登録**

`app.py` の blueprint import 群に追加:

```python
    from routes.ledger import bp as ledger_bp
```

register 群に追加:

```python
    app.register_blueprint(ledger_bp)
```

- [ ] **Step 5: テストが通ることを確認**

Run: `python -m pytest tests/test_ledger_routes.py -v`
Expected: `5 passed`

- [ ] **Step 6: 全テスト実行**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 7: コミット**

```powershell
git add routes/ledger.py app.py tests/test_ledger_routes.py
git commit -m "feat(zaiko): 手動足し引き・残高検索・履歴のAPIを追加"
```

---

### Task 8: 台帳画面（検索・足し引き・履歴UI）

**Files:**
- Create: `templates/ledger.html`
- Modify: `templates/base.html:31` (ナビリンク追加)

**Interfaces:**
- Consumes: `GET /ledger/api/balances` / `POST /ledger/api/tx` / `GET /ledger/api/history/<code>` (Task 7)

- [ ] **Step 1: テンプレートを作成**

`templates/ledger.html`:

```html
{% extends 'base.html' %}

{% block content %}
<div class="topbar">
  <span class="page-title">在庫台帳（足し引き）</span>
  <div class="topbar-actions">
    <button class="btn btn-info btn-sm" onclick="loadBalances()">⟳ 更新</button>
  </div>
</div>

<div class="filter-bar" style="gap:12px;padding:8px 20px;">
  <input type="text" id="q" placeholder="商品コード / 商品名で検索"
         style="width:280px;font-size:13px;padding:6px 10px;"
         onkeydown="if(event.key==='Enter')loadBalances()">
  <button class="btn btn-primary btn-sm" onclick="loadBalances()">検索</button>
</div>

<div style="padding:0 20px 20px;">
  <table class="data-table" style="width:100%;">
    <thead>
      <tr>
        <th>商品コード</th><th>商品名</th><th style="text-align:right;">台帳在庫</th>
        <th style="text-align:right;">引当済</th><th style="text-align:right;">出品可能</th>
        <th>棚番号</th><th>操作</th>
      </tr>
    </thead>
    <tbody id="rows"><tr><td colspan="7">検索してください</td></tr></tbody>
  </table>
</div>

<!-- 足し引きモーダル -->
<div id="tx-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;">
  <div style="background:var(--bg-primary,#fff);max-width:420px;margin:10vh auto;padding:20px;border-radius:8px;">
    <h3 id="tx-title" style="margin:0 0 12px;">足し引き</h3>
    <div style="margin-bottom:10px;">
      <label>種別</label>
      <select id="tx-type" style="width:100%;padding:6px;">
        <option value="manual_in">＋ 手動入庫</option>
        <option value="manual_out">− 手動出庫</option>
        <option value="adjust">± 調整（実数訂正）</option>
      </select>
    </div>
    <div style="margin-bottom:10px;">
      <label>数量（調整のみマイナス可）</label>
      <input type="number" id="tx-qty" value="1" style="width:100%;padding:6px;">
    </div>
    <div style="margin-bottom:14px;">
      <label>理由（出庫・調整は必須）</label>
      <input type="text" id="tx-reason" placeholder="例: 破損 / 棚卸差異" style="width:100%;padding:6px;">
    </div>
    <div style="display:flex;gap:8px;justify-content:flex-end;">
      <button class="btn btn-sm" onclick="closeTxModal()">キャンセル</button>
      <button class="btn btn-primary btn-sm" onclick="submitTx()">記録する</button>
    </div>
  </div>
</div>

<!-- 履歴モーダル -->
<div id="hist-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;">
  <div style="background:var(--bg-primary,#fff);max-width:640px;margin:8vh auto;padding:20px;border-radius:8px;max-height:75vh;overflow:auto;">
    <h3 id="hist-title" style="margin:0 0 12px;">履歴</h3>
    <table class="data-table" style="width:100%;">
      <thead><tr><th>日時</th><th>種別</th><th style="text-align:right;">数量</th><th>理由</th></tr></thead>
      <tbody id="hist-rows"></tbody>
    </table>
    <div style="text-align:right;margin-top:12px;">
      <button class="btn btn-sm" onclick="document.getElementById('hist-modal').style.display='none'">閉じる</button>
    </div>
  </div>
</div>

<script>
let currentCode = '';

async function loadBalances() {
  const q = document.getElementById('q').value.trim();
  const res = await fetch('/ledger/api/balances?q=' + encodeURIComponent(q));
  const data = await res.json();
  const tbody = document.getElementById('rows');
  if (!data.items.length) {
    tbody.innerHTML = '<tr><td colspan="7">該当なし</td></tr>';
    return;
  }
  tbody.innerHTML = data.items.map(it => `
    <tr>
      <td>${esc(it.product_code)}</td>
      <td>${esc(it.product_name)}</td>
      <td style="text-align:right;font-weight:600;">${it.ledger_qty}</td>
      <td style="text-align:right;">${it.reserved_qty}</td>
      <td style="text-align:right;">${it.available_qty}</td>
      <td>${esc(it.location)}</td>
      <td>
        <button class="btn btn-sm" onclick="openTxModal('${esc(it.product_code)}')">±</button>
        <button class="btn btn-sm" onclick="openHistory('${esc(it.product_code)}')">履歴</button>
      </td>
    </tr>`).join('');
}

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g,
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function openTxModal(code) {
  currentCode = code;
  document.getElementById('tx-title').textContent = code + ' の足し引き';
  document.getElementById('tx-qty').value = 1;
  document.getElementById('tx-reason').value = '';
  document.getElementById('tx-modal').style.display = 'block';
}
function closeTxModal() {
  document.getElementById('tx-modal').style.display = 'none';
}

async function submitTx() {
  const body = {
    product_code: currentCode,
    tx_type: document.getElementById('tx-type').value,
    qty: parseInt(document.getElementById('tx-qty').value, 10),
    reason: document.getElementById('tx-reason').value.trim(),
  };
  const res = await fetch('/ledger/api/tx', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!res.ok) { alert(data.message || 'エラー'); return; }
  closeTxModal();
  loadBalances();
}

async function openHistory(code) {
  const res = await fetch('/ledger/api/history/' + encodeURIComponent(code));
  const data = await res.json();
  document.getElementById('hist-title').textContent = code + ' の履歴';
  document.getElementById('hist-rows').innerHTML = data.items.map(t => `
    <tr>
      <td>${esc(t.created_at)}</td>
      <td>${esc(t.tx_type_label)}</td>
      <td style="text-align:right;font-weight:600;color:${t.qty >= 0 ? 'var(--success,#188038)' : 'var(--danger,#d93025)'};">
        ${t.qty >= 0 ? '+' : ''}${t.qty}</td>
      <td>${esc(t.reason)}</td>
    </tr>`).join('') || '<tr><td colspan="4">履歴なし</td></tr>';
  document.getElementById('hist-modal').style.display = 'block';
}

loadBalances();
</script>
{% endblock %}
```

- [ ] **Step 2: ナビリンクを追加**

`templates/base.html` の31行目(「在庫リアルタイム確認」リンク)の直後に追加:

```html
      <a href="{{ url_for('ledger.index') }}" class="nav-item {{ 'active' if active_page == 'ledger' }}">在庫台帳（足し引き）</a>
```

- [ ] **Step 3: ブラウザで動作確認**

`.claude/launch.json` のサーバーを起動し `/ledger/` を開く。確認項目:
1. 検索で即納在庫が一覧表示される
2. 「±」→手動入庫 3個 → 台帳在庫が +3 される
3. 「±」→手動出庫(理由なし) → 「出庫・調整には理由が必要です」エラー
4. 「履歴」→ 今の入庫が履歴に出る

Expected: 上記4点すべて動作。コンソールにJSエラーなし。

- [ ] **Step 4: 全テスト実行**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 5: コミット**

```powershell
git add templates/ledger.html templates/base.html
git commit -m "feat(zaiko): 在庫台帳画面を追加（検索・手動足し引き・履歴表示）"
```

---

### Task 9: 日次整合性チェック（台帳 vs キャッシュ）

**Files:**
- Modify: `services/stock_ledger.py` (関数追加)
- Modify: `routes/import_data.py:1163以降` (`run_all_imports_job` 末尾にフック) と `import_all` エンドポイント
- Test: `tests/test_integrity_check.py`

**Interfaces:**
- Consumes: `get_balance`, `recalc_inventory` (Task 3)
- Produces: `verify_cache_integrity() -> list[dict]` — 台帳に登場する全商品について Σ台帳 と 即納 Inventory.quantity を比較。不一致は recalc で修正し `Alert(alert_type='ledger_mismatch')` を作成(同一商品の未解決アラートは重複させない)。戻り値は `[{product_code, expected, actual}]`。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_integrity_check.py`:

```python
from models import db, Alert, Inventory
from services.stock_ledger import record_transaction, verify_cache_integrity


def test_mismatch_is_fixed_and_alerted(app):
    record_transaction('P-001', 'manual_in', 5, 'manual:i1')
    db.session.commit()

    # 台帳を通さない直接書き換え（あってはならない操作）を再現
    inv = Inventory.query.filter_by(product_code='P-001', inventory_type='即納').first()
    inv.quantity = 99
    db.session.commit()

    mismatches = verify_cache_integrity()
    db.session.commit()
    assert mismatches == [{'product_code': 'P-001', 'expected': 5, 'actual': 99}]
    assert inv.quantity == 5  # 修正済み
    assert Alert.query.filter_by(alert_type='ledger_mismatch').count() == 1

    # 再実行: 一致しているので何も起きない
    assert verify_cache_integrity() == []
    assert Alert.query.filter_by(alert_type='ledger_mismatch').count() == 1
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `python -m pytest tests/test_integrity_check.py -v`
Expected: FAIL — `ImportError: cannot import name 'verify_cache_integrity'`

- [ ] **Step 3: 実装**

`services/stock_ledger.py` 末尾に追加:

```python
def verify_cache_integrity():
    """台帳合計と即納Inventoryキャッシュの一致を検証し、不一致は修正+アラート。"""
    mismatches = []
    codes = [row[0] for row in
             db.session.query(StockTransaction.product_code).distinct().all()]
    for code in codes:
        expected = get_balance(code)
        inv = Inventory.query.filter_by(
            product_code=code, inventory_type='即納').first()
        actual = inv.quantity if inv else None
        if inv is not None and actual == expected:
            continue
        mismatches.append({'product_code': code,
                           'expected': expected, 'actual': actual})
        recalc_inventory(code)
        exists = Alert.query.filter_by(
            alert_type='ledger_mismatch', product_code=code,
            resolved_flag=False).first()
        if not exists:
            db.session.add(Alert(
                alert_type='ledger_mismatch',
                product_code=code,
                message=(f'{code}: 台帳合計 {expected} に対しキャッシュが {actual} '
                         f'でした。自動修正済み。台帳を通さない在庫変更が疑われます。'),
            ))
    return mismatches
```

- [ ] **Step 4: 自動実行にフック**

`routes/import_data.py` の `run_all_imports_job` の `with app.app_context():` ブロック末尾に追加:

```python
        # ─── 台帳整合性チェック ────────────────────────────────────────
        try:
            from services.stock_ledger import verify_cache_integrity
            mismatches = verify_cache_integrity()
            db.session.commit()
            if mismatches:
                app.logger.warning(f'Ledger integrity: {len(mismatches)} 件修正 {mismatches}')
        except Exception as e:
            db.session.rollback()
            app.logger.error(f'Ledger integrity check error: {e}')
```

`import_all` エンドポイント(803行付近)の最終 `return` の直前にも以下を追加する。VPSのsystemdタイマーは `import_all` を叩くため、両方に入れることで日次チェックが確実に走る:

```python
    # ─── 台帳整合性チェック ────────────────────────────────────────
    try:
        from services.stock_ledger import verify_cache_integrity
        mismatches = verify_cache_integrity()
        db.session.commit()
    except Exception:
        db.session.rollback()
        mismatches = []
```

(`import_all` のレスポンス JSON に `'ledger_mismatch_fixed': len(mismatches)` を1キー追加する)

- [ ] **Step 5: テストが通ることを確認**

Run: `python -m pytest tests -v`
Expected: 全件 PASS

- [ ] **Step 6: コミット**

```powershell
git add services/stock_ledger.py routes/import_data.py tests/test_integrity_check.py
git commit -m "feat(zaiko): 台帳とキャッシュの日次整合性チェックを追加（自動修正+アラート）"
```

---

### Task 10: 本番切替手順の実行(手動ステップ)

**Files:** なし(運用作業)

- [ ] **Step 1: ローカルDBでシード実行**

Run: `python scripts/seed_ledger_opening.py`
Expected: `期首残高: {'seeded': N, 'skipped': M} / YahooモールSKU: {'created': X, 'skipped': Y}` (N,X は環境の商品数に依存)

- [ ] **Step 2: 画面で期首残高を確認**

`/ledger/` で数商品を検索し、台帳在庫が従来の即納在庫数と一致することを確認。履歴に「期首残高」の adjust が1件ずつあること。

- [ ] **Step 3: 注文取込を1回実行して台帳連動を確認**

`/import/yahoo_orders` を実行(画面の取込ボタン or POST)。レスポンスの `ledger_out` / `ledger_returns` を確認し、新規注文分だけ台帳が減っていること・既存注文が二重に引かれていないことを `/ledger/` の履歴で確認。

Expected: 期首残高+新規注文出庫のみが履歴に並ぶ。

- [ ] **Step 4: VPSへのデプロイ**

VPS側の運用(git pull + サービス再起動 + `python scripts/seed_ledger_opening.py` 実行)。デプロイ手順の詳細は `deploy/systemd/` の README を参照。

---

## Self-Review 結果

- **スペックカバレッジ**: Phase 1 スコープ(台帳・冪等性・注文自動マイナス/キャンセル戻し・出荷済みキャンセルのアラート・手動足し引きUI・履歴画面・入庫連動・キャッシュ再計算・日次整合性チェック・mall_skus先行作成・期首シード)はTask 1-10で全て対応。書き戻し(Phase 2)・棚卸(Phase 3)・マルチモール(Phase 4)は本計画のスコープ外。
- **プレースホルダ**: なし(全ステップに実コード・実コマンド・期待結果を記載)。
- **型・名前の整合**: `record_transaction` のシグネチャはTask 3で定義しTask 4-9で同一使用。`source_key` 規約(`seed:` / `yahoo:{order}:{item}:out|return` / `japan_staging:{id}` / `manual:{uuid}`)は全タスクで統一。
