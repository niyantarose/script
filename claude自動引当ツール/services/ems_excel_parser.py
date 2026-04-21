"""韓国代行EMSファイル（Excel）のパーサー
ファイル名: EMS発送リスト_YYMMDD_HHMM.xlsx  /  EMS발송리스트_YYMMDD_HHMM.xlsx
シート: Sheet1

行1-3: 送付先住所情報（スキップ）
行4  : 韓国語ヘッダー行（スキップ）
行5以降: データ行（B列が 'Wata' で始まる行）

列マッピング（実際のファイルに基づく）:
  A(0) : #              → 行番号
  B(1) : 구매일         → 購買日（Wata形式）
  C(2) : 발송일         → 発送日（YYMMDD 形式: 250204 = 2025/02/04）
  D(3) : 입고수량       → 入荷数量
  E(4) : -              → 未使用
  F(5) : 구매번호       → 発注NO
  G(6) : 주문번호       → Yahoo受注番号
  H(7) : 주문번호2      → 未使用
  I(8) : 업체명         → 発注先
  J(9) : 상품내용       → 商品名
  K(10): 상품ID         → danielID（空の場合が多い）
  L(11): 냔타로즈ID     → 商品コード（主キー）
  M(12): 주문수량       → 数量
  N(13): 옵션명         → オプション
  O(14): 특이사항       → メモ
  ...
  R(17): 배송1          → EMS追跡番号（例: EG 028 313 542 KR）
"""
from datetime import date


class EmsExcelParser:

    SHEET_NAME = 'Sheet1'
    SKIP_ROWS  = 4   # 行1-4はアドレス情報・ヘッダー行

    COL_PURCHASE_DATE = 1   # B: 購買日（Wata形式、データ行の識別に使用）
    COL_SHIPPED_AT    = 2   # C: 発送日（YYMMDD形式）
    COL_PURCHASE_NO   = 5   # F: 発注NO
    COL_ORDER_ID      = 6   # G: Yahoo受注番号
    COL_SHOP_NAME     = 8   # I: 発注先
    COL_PRODUCT_NAME  = 9   # J: 商品名
    COL_PRODUCT_CODE  = 11  # L: 商品コード（냔타로즈ID）
    COL_QUANTITY      = 12  # M: 数量
    COL_MEMO          = 14  # O: メモ
    COL_EMS_NUMBER    = 17  # R: EMS追跡番号

    def __init__(self):
        pass

    # ──────────────────────────────────────────────
    def parse(self, file_path):
        """Excelファイルをパースして EMS アイテムリストを返す"""
        try:
            import openpyxl
        except ImportError:
            raise ImportError('openpyxl が必要です: pip install openpyxl')

        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

        if self.SHEET_NAME in wb.sheetnames:
            ws = wb[self.SHEET_NAME]
        else:
            ws = wb.active

        items = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
            if row_idx < self.SKIP_ROWS:
                continue

            cells = list(row)
            if not any(cells):
                continue

            # B列が 'Wata' で始まる行のみ取り込む
            purchase_date = self._cell(cells, self.COL_PURCHASE_DATE)
            if not purchase_date or not purchase_date.startswith('Wata'):
                continue

            # EMS番号（R列）は必須
            ems_raw = self._cell(cells, self.COL_EMS_NUMBER)
            ems_number = ems_raw.replace(' ', '').strip() if ems_raw else ''
            if not ems_number:
                continue

            # 商品コード: L列（냔타로즈ID）
            product_code = self._cell(cells, self.COL_PRODUCT_CODE)
            if not product_code:
                product_code = self._cell(cells, 10)  # K列フォールバック

            item = {
                'ems_number':    ems_number,
                'shipped_at':    self._parse_yymmdd(self._cell(cells, self.COL_SHIPPED_AT)),
                'purchase_date': purchase_date,                          # B列: 구매日 作成日
                'purchase_no':   self._cell(cells, self.COL_PURCHASE_NO),  # F列: 구매番号
                'order_id':      self._cell(cells, self.COL_ORDER_ID),
                'shop_name':     self._cell(cells, self.COL_SHOP_NAME),
                'product_name':  self._cell(cells, self.COL_PRODUCT_NAME),
                'product_code':  product_code,
                'quantity':      self._to_int(cells, self.COL_QUANTITY, default=1),
                'memo':          self._cell(cells, self.COL_MEMO),
            }
            items.append(item)

        wb.close()
        return items

    # ──────────────────────────────────────────────
    @staticmethod
    def _cell(cells, idx):
        if idx >= len(cells):
            return ''
        v = cells[idx]
        return str(v).strip() if v is not None else ''

    @staticmethod
    def _to_int(cells, idx, default=0):
        try:
            v = cells[idx] if idx < len(cells) else None
            return int(float(v)) if v else default
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _parse_yymmdd(raw):
        """
        '250204' → date(2025, 2, 4)
        '260108' → date(2026, 1, 8)
        失敗・空欄時は None を返す（today() デフォルトにしない）
        """
        try:
            s = str(raw).strip()
            # 数字6桁のみ抽出
            digits = ''.join(c for c in s if c.isdigit())[:6]
            if len(digits) == 6:
                yy = int(digits[0:2])
                mm = int(digits[2:4])
                dd = int(digits[4:6])
                return date(2000 + yy, mm, dd)
        except (ValueError, IndexError):
            pass
        return None   # ← today() デフォルトを廃止（sort汚染防止）
