// ============================================================
// 商品マスタ：新リスト（取込タブ）をマージ
//   設計: docs/superpowers/specs/2026-06-26-商品マスタ-新リストマージ-design.md
//   ・取込タブ「新マスタ取込」に新リスト(list)を貼り付け（ヘッダー名で列判定: NAME/CODE/ITEM TYPE/price）
//   ・被りコード: 品目・商品名・価格を新リストで上書き（空欄は上書きしない）。業者・重量は既存キープ
//   ・新規コード: 追加（業者・重量は空＝後で発注から補完）
//   ・旧マスタだけの行: 品目を英語化（「品目マッピング」タブ A=元/B=英語。未変換はA列へ自動追記）
//   ・実行前: ドライラン確認＋商品マスタ自動バックアップ＋実行ログ
//   依存: normCode_ / _masterPrimaryCodeKey_（既存）
// ============================================================

const MASTER_MERGE_CFG = {
  MASTER_SHEET: '商品マスタ',
  MASTER_HEADER_ROW: 6,   // データは7行目から。B:Gの6列が1レコード(コード/業者/商品名/品目/価格/重さ)
  M_CODE: 2,              // B列（1-based）。ここから6列を読み書き
  M_ITEM: 5,              // E列 品目（1-based）
  IMPORT_SHEET: '新マスタ取込',
  MAP_SHEET: '品目マッピング'
};

function _masterMergeKey_(code) {
  return (typeof _masterPrimaryCodeKey_ === 'function') ? _masterPrimaryCodeKey_(code) : normCode_(code);
}

// 品目の照合用キー（NFKC・空白除去・小文字）
function _品目キー_(v) {
  return String(v || '').normalize('NFKC').replace(/[\s　]+/g, '').trim().toLowerCase();
}

// 既に英語表記か（ASCIIのみ＝日本語/韓国語が無い）
function _品目は英語か_(v) {
  const s = String(v || '').trim();
  return s !== '' && /^[\x00-\x7F]+$/.test(s);
}

// 品目を英語化する。マッピング優先 → ASCIIはそのまま → それ以外は元のまま残し unmapped に記録。
function _品目を英語化_(value, itemMap, unmapped) {
  const cur = String(value || '').trim();
  if (!cur) return '';
  const mapped = itemMap[_品目キー_(cur)];
  if (mapped) return mapped;          // マッピングがあれば英語へ（ASCIIでも上書き可）
  if (_品目は英語か_(cur)) return cur; // 既に英語ならそのまま
  if (unmapped) unmapped[cur] = (unmapped[cur] || 0) + 1; // 未マッピングは元のまま＋記録
  return cur;
}

function 商品マスタ_新リストをマージ() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const t0 = new Date();
  const L = m => Logger.log('[新リストマージ] ' + m);
  L('開始 ' + Utilities.formatDate(t0, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));

  const master = ss.getSheetByName(MASTER_MERGE_CFG.MASTER_SHEET);
  const imp = ss.getSheetByName(MASTER_MERGE_CFG.IMPORT_SHEET);
  if (!master) { ui.alert('「' + MASTER_MERGE_CFG.MASTER_SHEET + '」が見つかりません。'); return; }
  if (!imp) { ui.alert('取込タブ「' + MASTER_MERGE_CFG.IMPORT_SHEET + '」が見つかりません。\n新リスト(list)の中身を貼り付けてタブを作ってください。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    // 1) 取込データ（ヘッダー名で列判定・コードはキーで先勝ち重複排除）
    const imported = 商品マスタ取込_読み込み_(imp, L);
    if (!imported.records.length) {
      ui.alert('取込タブにデータがありません。\nヘッダー行に NAME / CODE / ITEM TYPE / price があるか確認してください。');
      L('終了: 取込データなし'); return;
    }

    // 2) 商品マスタ B:G ブロック読み込み（[code,vendor,name,item,price,weight]）
    const mHeader = MASTER_MERGE_CFG.MASTER_HEADER_ROW;
    const mStart = mHeader + 1;
    const mRows = Math.max(0, master.getLastRow() - mHeader);
    const block = mRows > 0 ? master.getRange(mStart, MASTER_MERGE_CFG.M_CODE, mRows, 6).getValues() : [];
    const masterIndex = {};
    for (let i = 0; i < block.length; i++) {
      const k = _masterMergeKey_(block[i][0]);
      if (k && !(k in masterIndex)) masterIndex[k] = i;
    }

    // 3) 品目マッピング
    const itemMap = 商品マスタ_品目マッピング読み込み_(ss);

    // 4) 取込を適用算出（被り=上書き / 新規=追加）。品目はマッピング経由で英語化。
    const unmapped = {};
    const importKeys = {};
    let updated = 0, added = 0, unchanged = 0;
    const newRows = [];
    imported.records.forEach(r => {
      const k = _masterMergeKey_(r.code);
      importKeys[k] = true;
      if (k in masterIndex) {
        const row = block[masterIndex[k]];
        const before = [row[2], row[3], row[4]].join('│');
        if (String(r.name).trim() !== '') row[2] = r.name;                              // D 商品名
        if (String(r.item).trim() !== '') row[3] = _品目を英語化_(r.item, itemMap, unmapped); // E 品目（英語化）
        if (r.price !== '' && r.price !== null) row[4] = r.price;                       // F 価格
        if ([row[2], row[3], row[4]].join('│') !== before) updated++; else unchanged++;
      } else {
        const newItem = _品目を英語化_(r.item, itemMap, unmapped);
        newRows.push([r.code, '', r.name, newItem, r.price, '']);          // B/D/E/F、業者C・重量Gは空
        masterIndex[k] = block.length + newRows.length - 1;                // 同一実行内の二重追加防止
        added++;
      }
    });

    // 5) 旧マスタだけの行の品目を英語化
    let itemConverted = 0;
    for (let i = 0; i < block.length; i++) {
      const k = _masterMergeKey_(block[i][0]);
      if (importKeys[k]) continue;                 // 取込にある行は上で処理済み
      const cur = String(block[i][3] || '').trim();
      if (!cur) continue;
      const to = _品目を英語化_(cur, itemMap, unmapped);
      if (to !== cur) { block[i][3] = to; itemConverted++; }
    }
    const unmappedList = Object.keys(unmapped);
    if (unmappedList.length) 商品マスタ_未マッピングを品目マッピングへ追記_(ss, unmappedList);

    L('算出: 更新' + updated + ' / 追加' + added + ' / 変更なし' + unchanged +
      ' / 旧行品目英語化' + itemConverted + ' / 取込重複' + imported.dupCount + ' / 未マッピング' + unmappedList.length);

    // ドライラン確認
    const dupNote = imported.dupCount ? '\n（取込内の重複コードは先勝ちで' + imported.dupCount + '件スキップ）' : '';
    const unmapNote = unmappedList.length
      ? '\n\n⚠ 未マッピング品目 ' + unmappedList.length + '件 →「品目マッピング」タブのA列に追記しました。\n  B列に英語を入れて再実行すると英語化されます。\n  ' +
        unmappedList.slice(0, 15).join(', ') + (unmappedList.length > 15 ? ' …ほか' : '')
      : '';
    const res = ui.alert('新リスト → 商品マスタ マージ（確認）',
      '商品マスタへマージします。\n\n' +
      '■更新（被りコード）: ' + updated + '件\n' +
      '■新規追加: ' + added + '件\n' +
      '■変更なし: ' + unchanged + '件\n' +
      '■旧行の品目を英語化: ' + itemConverted + '件' + dupNote + unmapNote +
      '\n\n実行前に商品マスタを自動バックアップします。実行しますか？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) {
      L('中止: ユーザーキャンセル（品目マッピングへの追記のみ反映）');
      ui.alert('やめました。' + (unmappedList.length ? '\n「品目マッピング」のB列を埋めてから再実行してください。' : ''));
      return;
    }

    // バックアップ
    const bkName = '商品マスタ_backup_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
    master.copyTo(ss).setName(bkName);
    L('バックアップ作成: ' + bkName);

    // 適用（既存ブロック上書き＋新規追記）
    if (block.length) master.getRange(mStart, MASTER_MERGE_CFG.M_CODE, block.length, 6).setValues(block);
    if (newRows.length) master.getRange(mStart + block.length, MASTER_MERGE_CFG.M_CODE, newRows.length, 6).setValues(newRows);
    SpreadsheetApp.flush();

    L('完了: 追加' + added + ' / 更新' + updated + ' / 品目英語化' + itemConverted +
      ' / バックアップ=' + bkName + ' / 所要' + (new Date() - t0) + 'ms');
    ss.toast('マージ完了: 追加' + added + ' / 更新' + updated + ' / 品目英語化' + itemConverted);
    ui.alert('マージ完了',
      '追加: ' + added + '件\n更新: ' + updated + '件\n旧行の品目英語化: ' + itemConverted + '件\n' +
      'バックアップ: ' + bkName +
      (unmappedList.length ? '\n\n⚠ 未マッピング品目が' + unmappedList.length + '件残っています。\n「品目マッピング」のB列を埋めて再実行すると英語化されます。' : ''),
      ui.ButtonSet.OK);
  } catch (err) {
    L('🛑 例外: ' + (err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// 取込タブを読み、{code,name,item,price} 配列を作る。ヘッダー名で列特定、コードキー先勝ちで重複排除。
function 商品マスタ取込_読み込み_(sh, L) {
  const log = L || function () {};
  const vals = sh.getDataRange().getValues();
  let h = -1, cName = -1, cCode = -1, cItem = -1, cPrice = -1;
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i].map(v => String(v || '').trim());
    const ni = row.findIndex(c => c.toUpperCase() === 'NAME');
    const ci = row.findIndex(c => c.toUpperCase() === 'CODE');
    if (ni >= 0 && ci >= 0) {
      h = i; cName = ni; cCode = ci;
      cItem = row.findIndex(c => c.toUpperCase() === 'ITEM TYPE');
      cPrice = row.findIndex(c => c.toLowerCase() === 'price');
      break;
    }
  }
  if (h < 0) { log('取込: ヘッダー行(NAME/CODE)が見つかりません'); return { records: [], dupCount: 0 }; }

  const seen = {}; const records = []; let dupCount = 0;
  for (let i = h + 1; i < vals.length; i++) {
    const code = String(vals[i][cCode] || '').trim();
    if (!code) continue;
    const k = _masterMergeKey_(code);
    if (!k) continue;
    if (seen[k]) { dupCount++; log('取込重複(先勝ちスキップ): ' + code + ' / ' + String(vals[i][cName] || '').trim()); continue; }
    seen[k] = true;
    records.push({
      code: code,
      name: String(vals[i][cName] || '').trim(),
      item: cItem >= 0 ? String(vals[i][cItem] || '').trim() : '',
      price: cPrice >= 0 ? vals[i][cPrice] : ''
    });
  }
  log('取込読み込み: ' + records.length + '件（重複スキップ ' + dupCount + '件）');
  return { records: records, dupCount: dupCount };
}

// 品目マッピング {元品目キー: 英語}
function 商品マスタ_品目マッピング読み込み_(ss) {
  const sh = ss.getSheetByName(MASTER_MERGE_CFG.MAP_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 1) return map;
  const vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    const from = String(vals[i][0] || '').trim();
    const to = String(vals[i][1] || '').trim();
    if (!from || !to) continue;
    if (/^(元|品目|from)/i.test(from)) continue; // 見出し行
    map[_品目キー_(from)] = to;
  }
  return map;
}

// 未マッピング品目を「品目マッピング」タブのA列へ追記（B列は空のまま）。無ければタブ作成。
function 商品マスタ_未マッピングを品目マッピングへ追記_(ss, list) {
  let sh = ss.getSheetByName(MASTER_MERGE_CFG.MAP_SHEET);
  if (!sh) { sh = ss.insertSheet(MASTER_MERGE_CFG.MAP_SHEET); sh.getRange(1, 1, 1, 2).setValues([['元の品目', '英語品目']]); }
  const last = sh.getLastRow();
  const existing = {};
  if (last >= 1) {
    sh.getRange(1, 1, last, 1).getValues().forEach(r => {
      const v = String(r[0] || '').trim();
      if (v) existing[_品目キー_(v)] = true;
    });
  }
  const toAdd = list.filter(v => !existing[_品目キー_(v)]).map(v => [v, '']);
  if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
}

// ============================================================
// 商品マスタ：品目(E列)を「品目マッピング」で英語に統一する（マージとは独立して実行可）
//   ・マッピングに有る→英語へ ／ 既に英語(ASCII)→そのまま ／ それ以外→元のまま残し未マッピング報告(A列へ追記)
//   ・ドライラン確認＋自動バックアップ＋ログ
// ============================================================
function 商品マスタ_品目を英語に統一() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const t0 = new Date();
  const L = m => Logger.log('[品目英語統一] ' + m);
  L('開始 ' + Utilities.formatDate(t0, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));

  const master = ss.getSheetByName(MASTER_MERGE_CFG.MASTER_SHEET);
  if (!master) { ui.alert('「' + MASTER_MERGE_CFG.MASTER_SHEET + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const itemMap = 商品マスタ_品目マッピング読み込み_(ss);
    const mStart = MASTER_MERGE_CFG.MASTER_HEADER_ROW + 1;
    const n = Math.max(0, master.getLastRow() - MASTER_MERGE_CFG.MASTER_HEADER_ROW);
    if (n <= 0) { ui.alert('商品マスタにデータがありません。'); return; }

    const range = master.getRange(mStart, MASTER_MERGE_CFG.M_ITEM, n, 1); // E列 品目
    const vals = range.getValues();
    const unmapped = {};
    let converted = 0, alreadyEn = 0, empty = 0;
    for (let i = 0; i < n; i++) {
      const cur = String(vals[i][0] || '').trim();
      if (!cur) { empty++; continue; }
      const to = _品目を英語化_(cur, itemMap, unmapped);
      if (to !== cur) { vals[i][0] = to; converted++; }
      else if (_品目は英語か_(cur) || itemMap[_品目キー_(cur)]) alreadyEn++;
    }
    const unmappedList = Object.keys(unmapped);
    if (unmappedList.length) 商品マスタ_未マッピングを品目マッピングへ追記_(ss, unmappedList);

    L('算出: 英語へ変換' + converted + ' / 既に英語' + alreadyEn + ' / 空' + empty + ' / 未マッピング' + unmappedList.length);

    const unmapNote = unmappedList.length
      ? '\n\n⚠ 未マッピング品目 ' + unmappedList.length + '件 →「品目マッピング」A列に追記しました。\n  B列に英語を入れて再実行で変換されます。\n  ' +
        unmappedList.slice(0, 20).join(', ') + (unmappedList.length > 20 ? ' …ほか' : '')
      : '';
    const res = ui.alert('商品マスタ：品目を英語に統一（確認）',
      '商品マスタの品目(E列)を英語に統一します。\n\n' +
      '■英語へ変換: ' + converted + '件\n' +
      '■既に英語: ' + alreadyEn + '件\n' +
      '■空欄: ' + empty + '件' + unmapNote +
      '\n\n実行前に商品マスタを自動バックアップします。実行しますか？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) {
      L('中止: ユーザーキャンセル（品目マッピングへの追記のみ反映）');
      ui.alert('やめました。' + (unmappedList.length ? '\n「品目マッピング」のB列を埋めてから再実行してください。' : ''));
      return;
    }

    if (converted > 0) {
      const bkName = '商品マスタ_backup_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
      master.copyTo(ss).setName(bkName);
      L('バックアップ作成: ' + bkName);
      range.setValues(vals);
      SpreadsheetApp.flush();
      L('適用完了: ' + converted + '件変換 / 所要' + (new Date() - t0) + 'ms');
      ss.toast('品目を英語化: ' + converted + '件（未マッピング' + unmappedList.length + '件）');
    } else {
      L('変換対象なし');
      ss.toast('変換対象なし（未マッピング' + unmappedList.length + '件）');
    }
    ui.alert('完了',
      '英語へ変換: ' + converted + '件\n既に英語: ' + alreadyEn + '件' +
      (unmappedList.length
        ? '\n\n⚠ 未マッピング品目が' + unmappedList.length + '件残っています。\n「品目マッピング」のB列を埋めて再実行してください。'
        : '\n\nすべて英語に統一されました。'),
      ui.ButtonSet.OK);
  } catch (err) {
    L('🛑 例外: ' + (err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}
