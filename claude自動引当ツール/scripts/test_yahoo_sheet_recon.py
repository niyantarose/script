#!/usr/bin/env python3
"""yahoo_sheet_recon の純粋ロジック検証（ネットワーク不要）。
実行: python scripts/test_yahoo_sheet_recon.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import yahoo_sheet_recon as r

failed = 0


def eq(name, actual, expected):
    global failed
    if actual != expected:
        failed += 1
        print(f'[NG] {name}: expected={expected!r} actual={actual!r}')
    else:
        print(f'[OK] {name}')


# ---- コード正規化: 全角/小文字/ダッシュ揺れを吸収（Yahoo側は小文字になる）----
eq('正規化: 全角→半角+大文字', r.normalize_code('ｔｗｆ００２７－ｃｍ－０６'), 'TWF0027-CM-06')
eq('正規化: 小文字+特殊ダッシュ', r.normalize_code(' tw0025‐cm‐06 '), 'TW0025-CM-06')
eq('正規化: 空', r.normalize_code('  '), '')

# ---- 照合②の対象スコープ: 台湾CN系コード形状のみ ----
eq('スコープ: TW書籍', bool(r.TAIWAN_CODE_RE.match('TW0001-CM-01')), True)
eq('スコープ: TWF特装', bool(r.TAIWAN_CODE_RE.match('TWF0027-CM-06')), True)
eq('スコープ: TWSセット', bool(r.TAIWAN_CODE_RE.match('TWS0079-CM-0102')), True)
eq('スコープ: CN', bool(r.TAIWAN_CODE_RE.match('CN0100-NV-01')), True)
eq('スコープ外: KR', bool(r.TAIWAN_CODE_RE.match('KR0001-CM-01')), False)
eq('スコープ外: 一般コード', bool(r.TAIWAN_CODE_RE.match('goq-12345')), False)

# ---- 突き合わせ本体 ----
sheet_rows = [
    # ①: 登録済みなのにYahooに無い
    {'sheet': '台湾まんが', 'row': 5, 'code': 'TW0010-CM-01', 'title': '作品A 1巻', 'status': '登録済み'},
    # 未登録行はYahooに無くても①に出さない
    {'sheet': '台湾まんが', 'row': 6, 'code': 'TW0011-CM-01', 'title': '作品B 1巻', 'status': ''},
    # ③: コード一致・タイトル不一致
    {'sheet': '台湾まんが', 'row': 7, 'code': 'TW0012-CM-01', 'title': 'ファントムバスターズ 1巻', 'status': '登録済み'},
    # 一致（タイトルもほぼ同じ）→どこにも出ない
    {'sheet': '台湾書籍その他', 'row': 8, 'code': 'TW0013-NV-01', 'title': '小説C 1巻', 'status': '登録済み'},
    # ④: シート内コード重複
    {'sheet': '台湾まんが', 'row': 9, 'code': 'TW0014-CM-01', 'title': '作品D 1巻', 'status': '登録済み'},
    {'sheet': '台湾まんが', 'row': 10, 'code': 'TW0014-CM-01', 'title': '作品E 1巻', 'status': '登録済み'},
]
yahoo_items = [
    {'code': 'tw0012-cm-01', 'name': '恋せよまやかし天使ども 1巻'},   # ③（タイトル別作品）
    {'code': 'tw0013-nv-01', 'name': '小説C 1巻'},                    # 一致
    {'code': 'tw0014-cm-01', 'name': '作品D 1巻'},                    # ④の片割れは店に存在
    {'code': 'tw9999-cm-01', 'name': '旧システムの何か'},              # ②
    {'code': 'kr0001-cm-01', 'name': '韓国の何か'},                    # スコープ外→②に出さない
]

buckets = r.reconcile(sheet_rows, yahoo_items)

eq('①件数', [x['code'] for x in buckets['sheet_only']], ['TW0010-CM-01'])
eq('②件数(スコープ内のみ)', [x['code'] for x in buckets['yahoo_only']], ['TW9999-CM-01'])
eq('③タイトル不一致を検出', [x['code'] for x in buckets['title_mismatch']], ['TW0012-CM-01'])
eq('③一致タイトルは出さない',
   any(x['code'] == 'TW0013-NV-01' for x in buckets['title_mismatch']), False)
eq('④シート内重複', [x['code'] for x in buckets['dup_in_sheet']], ['TW0014-CM-01'])

# ---- 差分検出（定期実行で「新規発生」だけを浮かせる）----
rows_now = r.buckets_to_key_rows(buckets)
prev_all = {(k, c) for k, c, _ in rows_now}
eq('差分: 前回と同じなら新規なし', r.find_new_findings(prev_all, buckets), [])
prev_minus = prev_all - {('②', 'TW9999-CM-01')}
news = r.find_new_findings(prev_minus, buckets)
eq('差分: 新規1件を検出', [(k, c) for k, c, _ in news], [('②', 'TW9999-CM-01')])
eq('差分: 初回(前回なし)は全件baseline扱い', r.find_new_findings(None, buckets), [])

# 大文字小文字・全角が違ってもコードは同一視される
buckets2 = r.reconcile(
    [{'sheet': 's', 'row': 2, 'code': 'ＴＷ００２０－ＣＭ－０１', 'title': 'X', 'status': '登録済み'}],
    [{'code': 'tw0020-cm-01', 'name': 'X'}],
)
eq('正規化マッチ: ①に出ない', buckets2['sheet_only'], [])

print(f'\n{failed} failed' if failed else '\nall passed')
sys.exit(1 if failed else 0)
