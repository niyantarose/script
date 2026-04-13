"""韓国代行EMSファイル（Excel）のパーサー"""
import os
from datetime import datetime


class EmsExcelParser:
    """
    韓国代行から届くEMSファイル（Excel形式）をパースする。
    列構成は別途共有予定のため、設定ファイルで列マッピングを管理する。
    """

    # デフォルトの列マッピング（別途共有されたら更新する）
    DEFAULT_COLUMN_MAP = {
        'ems_number': 0,       # EMS追跡番号
        'shipped_at': 1,       # 発送日
        'product_code': 2,     # 商品コード
        'product_name': 3,     # 商品名
        'quantity': 4,         # 数量
        'order_id': 5,         # 受注番号
    }

    def __init__(self, column_map=None):
        self.column_map = column_map or self.DEFAULT_COLUMN_MAP

    def parse(self, file_path):
        """Excelファイルをパースしてリストで返す"""
        try:
            import openpyxl
        except ImportError:
            raise ImportError('openpyxl が必要です: pip install openpyxl')

        wb = openpyxl.load_workbook(file_path, read_only=True)
        ws = wb.active

        items = []
        rows = list(ws.rows)
        if len(rows) < 2:
            return items

        # 1行目はヘッダーとしてスキップ
        for row in rows[1:]:
            cells = [cell.value for cell in row]
            if not any(cells):
                continue

            item = {}
            for key, col_idx in self.column_map.items():
                if col_idx < len(cells):
                    val = cells[col_idx]
                    if key == 'shipped_at' and isinstance(val, datetime):
                        val = val.date()
                    elif key == 'quantity':
                        val = int(val) if val else 0
                    else:
                        val = str(val).strip() if val else ''
                    item[key] = val
                else:
                    item[key] = ''

            if item.get('ems_number'):
                items.append(item)

        wb.close()
        return items
