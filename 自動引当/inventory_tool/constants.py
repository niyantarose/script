ORDER_ITEM_STATUS_LABELS = {
    "pending": "引当待ち",
    "provisional_allocated": "お取り寄せ引当済み（仮）",
    "allocated_sokunou": "即納引当済",
    "partial_waiting": "部分引当中（EMS待ち）",
    "priority_hold": "先送り待機",
    "shortage": "在庫不足（未引当）",
    "fully_allocated": "引当完了",
    "shipped": "発送完了",
}

ORDER_STATUS_LABELS = {
    "pending": "対応中",
    "allocated": "発送可能",
    "shipped": "発送済み",
}

ALERT_TYPE_LABELS = {
    "purchase_missing": "発注漏れ",
    "korea_ship_missing": "韓国発送漏れ",
    "japan_arrival_missing": "日本入荷漏れ",
    "japan_ship_missing": "発送漏れ",
    "stock_shortage": "在庫不足",
    "delay_warning": "遅延警告",
}

DELAY_LEVEL_LABELS = {
    "danger": "赤",
    "warning": "オレンジ",
    "notice": "黄",
    "ok": "緑",
}

JAPAN_STAGING_STATUS_LABELS = {
    "waiting": "仕分け待ち",
    "assigned_to_order": "受注に割当済み",
    "to_japan_stock": "日本在庫反映待ち",
    "excluded": "除外",
    "returned_to_ems": "EMS差し戻し",
    "reflected": "反映完了",
}

INVENTORY_TYPE_LABELS = {
    "即納": "即納",
    "お取り寄せ": "お取り寄せ",
}
