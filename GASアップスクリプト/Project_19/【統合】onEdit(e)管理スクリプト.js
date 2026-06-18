function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const range = e.range;

  try {
    // =====================================================
    // 発注シート側の処理
    // =====================================================
    if (sheetName === HATCHU_CFG.SHEET_NAME) {
      const editStartCol = range.getColumn();
      const editEndCol = range.getLastColumn();
      const editStartRow = range.getRow();
      const editNumRows = range.getNumRows();

      if (
        _rangeHitsCol_(editStartCol, editEndCol, CFG.HACHU_CODE) ||
        _rangeHitsCol_(editStartCol, editEndCol, CFG.HACHU_VENDOR)
      ) {
        autofillHachu(e);
      }

      商品コードと業者と商品名と価格がそろったら自動でマスタ登録(e);

      const onlyCheckboxCol =
        editStartCol === HATCHU_CFG.COL_CHK &&
        editEndCol === HATCHU_CFG.COL_CHK;

      if (!onlyCheckboxCol) {
        const startRow = Math.max(HATCHU_CFG.START_ROW, editStartRow - 2);
        const endRow = Math.min(
          sh.getMaxRows(),
          editStartRow + editNumRows + 2
        );
        const numRows = endRow - startRow + 1;

        if (numRows > 0) {
          HATCHU_processBordersAndCheckboxes_(sh, startRow, numRows);
        }
      }

      if (typeof applyKeshikomiColorForEditedRows_ === 'function') {
        SpreadsheetApp.flush();
        applyKeshikomiColorForEditedRows_(sh, editStartRow, editNumRows);
      }

      if (_rangeHitsCol_(editStartCol, editEndCol, HATCHU_GROUP_COL)) {
        発注_drawGroupBorders_(sh);
      }

      return;
    }

    // =====================================================
    // EMSリスト側の処理
    // =====================================================
    if (typeof EMS_isTargetSheet_ !== 'function') return;
    if (typeof EMS_CFG === 'undefined') return;
    if (!EMS_isTargetSheet_(sh)) return;

    const hitDate = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_DATE);
    const hitEms = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_NO);

    const isEmsList =
      typeof KESHIKOMI_COLOR_CFG !== 'undefined' &&
      sheetName === KESHIKOMI_COLOR_CFG.EMS_SHEET_NAME;

    if (!hitDate && !hitEms) {
      if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
  SpreadsheetApp.flush();

  const orderSheet = e.source.getSheetByName(KESHIKOMI_COLOR_CFG.SHEET_NAME);
  if (orderSheet) {
    colorKeshikomiAllRows_(orderSheet);
  }
}
      return;
    }

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(1000)) return;

    try {
      if (hitDate) {
        EMS_updateStatusOnlyForEditedRows_(sh, range);
      }

      if (
        EMS_CFG.AUTO_BOX_ON_EDIT &&
        range.getNumRows() <= EMS_CFG.ONEDIT_MAX_ROWS
      ) {
        EMS_updateDatesByRows_(
          sh,
          range.getRow(),
          range.getNumRows(),
          false
        );
      } else if (EMS_CFG.AUTO_BOX_ON_EDIT) {
        SpreadsheetApp.getActive().toast(
          '大量編集のため、ボタンから全体更新を実行してください。'
        );
      }

      if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
        SpreadsheetApp.flush();

        const orderSheet = e.source.getSheetByName(KESHIKOMI_COLOR_CFG.SHEET_NAME);
        if (orderSheet) {
          colorKeshikomiAllRows_(orderSheet);
        }
      }

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    SpreadsheetApp.getActive().toast(
      'onEditエラー: ' + err.message,
      'エラー',
      8
    );
    throw err;
  }
}
function _rangeHitsCol_(startCol, endCol, targetCol) {
  targetCol = Number(targetCol);
  return !!targetCol && startCol <= targetCol && targetCol <= endCol;
}
// =====================================================
// 発注シート：チェックボックスと罫線の一括高速処理ロジック
// =====================================================
function HATCHU_processBordersAndCheckboxes_(sh, startRow, numRows) {
  const cfg = HATCHU_CFG;

  const maxCol = Number(cfg.MAX_COL || cfg.BORDER_COLS || 28);
  const colNo = Number(cfg.COL_NO || 2);
  const colChk = Number(cfg.COL_CHK || 1);

  startRow = Number(startRow);
  numRows = Number(numRows);

  if (!sh) return;
  if (!startRow || !numRows || numRows <= 0) return;

  const maxRows = sh.getMaxRows();

  // 編集行の前後も広めに見る
  startRow = Math.max(cfg.START_ROW, startRow - 2);
  const endRow = Math.min(maxRows, startRow + numRows + 4);
  numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  const noVals = sh
    .getRange(startRow, colNo, numRows, 1)
    .getDisplayValues();

  const checkboxRule = SpreadsheetApp
    .newDataValidation()
    .requireCheckbox()
    .build();

  const fullRange = sh.getRange(startRow, 1, numRows, maxCol);

  // まず対象範囲の罫線を全部消す
  fullRange.setBorder(
    false,
    false,
    false,
    false,
    false,
    false
  );

  // A列チェックボックスも対象範囲で一度整理
  const chkRangeAll = sh.getRange(startRow, colChk, numRows, 1);
  chkRangeAll.clearDataValidations();

  // No.がない行のチェックボックス内容だけ消す
  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const hasNo = String(noVals[i][0] || '').trim() !== '';

    const chkCell = sh.getRange(rowNo, colChk, 1, 1);

    if (hasNo) {
      chkCell.setDataValidation(checkboxRule);
    } else {
      chkCell.clearContent();
    }
  }

  // ここが重要：
  // No.がある行を「1行ずつ」格子罫線にする
  // これで最終データ行の下罫線も確実に入る
  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const hasNo = String(noVals[i][0] || '').trim() !== '';

    if (!hasNo) continue;

    sh.getRange(rowNo, 1, 1, maxCol)
      .setBorder(
        true,   // 上
        true,   // 左
        true,   // 下
        true,   // 右
        true,   // 縦線
        true,   // 横線
        '#000000',
        SpreadsheetApp.BorderStyle.SOLID
      );
  }
}

/*** 設定:自分の列に合わせる(A=1,B=2,…)。商品マスタの列はうろ覚えやから要確認 ***/
/*** 設定:自分の列に合わせる(A=1,B=2,…) ***/
const CFG = {
  HACHU_SHEET: '発注',
  HACHU_HEADER_ROW: 6,

  // 発注シート側
  HACHU_CODE: 12,    // L列 商品コード
  HACHU_VENDOR: 9,   // I列 業者
  HACHU_NAME: 10,    // J列 商品名
  HACHU_ITEM: 15,    // O列 品目
  HACHU_WEIGHT: 16,  // P列 重さ
  HACHU_PRICE: 17,   // Q列 価格

  // 商品マスタ側
  MASTER_SHEET: '商品マスタ',
  MASTER_HEADER_ROW: 6,

  M_NO: 0,
  M_CODE: 2,     // B列 商品コード
  M_VENDOR: 3,   // C列 業者
  M_NAME: 4,     // D列 商品名
  M_ITEM: 5,     // E列 品目
  M_PRICE: 6,    // F列 価格
  M_WEIGHT: 7,   // G列 重さ

  UPDATE_PRICE_IF_DIFF: true,
  SEP: '│'
};

/************************************************************
 * 商品マスタ 既存データ補完設定
 * 今は「重さ」だけ。
 * 今後列が増えたら MASTER_SYNC_FIELD_RULES に1つ追加するだけでOK。
 ************************************************************/
const MASTER_SYNC_FIELD_RULES = [
  {
    label: '重さ',
    sourceCfgKey: 'HACHU_WEIGHT',
    sourceHeaders: ['重さ', 'weight(g)', 'weight', '重量'],
    masterCfgKey: 'M_WEIGHT',
    masterHeaders: ['重さ', 'weight(g)', 'weight', '重量'],
    overwrite: true // true=既存値も違えば更新 / false=空欄だけ補完
  }

  // 例：今後「サイズ」列も同期したくなったらこう追加
  // ,
  // {
  //   label: 'サイズ',
  //   sourceCfgKey: 'HACHU_SIZE',
  //   sourceHeaders: ['サイズ', 'Size'],
  //   masterCfgKey: 'M_SIZE',
  //   masterHeaders: ['サイズ', 'Size'],
  //   overwrite: true
  // }
];


/**
 * 発注シートの既存データから商品マスタを一括補完・更新する
 * キー：商品コード + 業者
 */
function 商品マスタ_既存データを発注から補完更新(silent) {
  const ss = SpreadsheetApp.getActive();
  const cfg = (typeof CFG !== 'undefined') ? CFG : {};

  const hachu = ss.getSheetByName(cfg.HACHU_SHEET || '発注');
  const master = ss.getSheetByName(cfg.MASTER_SHEET || '商品マスタ');
  // トリガー実行時はUIが使えないのでログに切り替える
  const say = msg => { if (silent) Logger.log(msg); else SpreadsheetApp.getUi().alert(msg); };

  if (!hachu) {
    say('発注シートが見つからんで');
    return;
  }
  if (!master) {
    say('商品マスタシートが見つからんで');
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    say('今ほかの処理が動いてるみたい。少し待ってもう一回。');
    return;
  }

  try {
    const hHeaderRow = cfg.HACHU_HEADER_ROW || 6;
    const mHeaderRow = cfg.MASTER_HEADER_ROW || 1;

    const hBase = hHeaderRow + 1;
    const mBase = mHeaderRow + 1;

    const hLastRow = hachu.getLastRow();
    const mLastRow = master.getLastRow();

    if (hLastRow < hBase) {
      say('発注データが無いで');
      return;
    }
    if (mLastRow < mBase) {
      say('商品マスタにデータが無いで');
      return;
    }

    const hHeaderRows = _補完_headerRows_(hHeaderRow);
    const mHeaderRows = _補完_headerRows_(mHeaderRow);

    const hCodeCol = _補完_resolveCol_(hachu, 'HACHU_CODE', ['商品コード', 'Code', 'code'], hHeaderRows);
    const hVendorCol = _補完_resolveCol_(hachu, 'HACHU_VENDOR', ['業者', 'Vendor', 'vendor'], hHeaderRows);

    const mCodeCol = _補完_resolveCol_(master, 'M_CODE', ['商品コード', 'Code', 'code'], mHeaderRows);
    const mVendorCol = _補完_resolveCol_(master, 'M_VENDOR', ['業者', 'Vendor', 'vendor'], mHeaderRows);

    if (!hCodeCol || !hVendorCol || !mCodeCol || !mVendorCol) {
      say(
        '商品コード列または業者列が見つからんで。\n\n' +
        'CFGの HACHU_CODE / HACHU_VENDOR / M_CODE / M_VENDOR を確認してな。'
      );
      return;
    }

    const rules = MASTER_SYNC_FIELD_RULES.map(rule => {
      return {
        label: rule.label,
        sourceCol: _補完_resolveCol_(hachu, rule.sourceCfgKey, rule.sourceHeaders, hHeaderRows),
        masterCol: _補完_resolveCol_(master, rule.masterCfgKey, rule.masterHeaders, mHeaderRows),
        overwrite: rule.overwrite !== false
      };
    });

    const usableRules = rules.filter(r => r.sourceCol && r.masterCol);
    const missingRules = rules.filter(r => !r.sourceCol || !r.masterCol);

    if (usableRules.length === 0) {
      say('補完対象の列が見つからんで。重さ列の見出し、またはCFGの M_WEIGHT / HACHU_WEIGHT を確認してな。');
      return;
    }

    const sep = cfg.SEP || '___SEP___';

    const hLastCol = hachu.getLastColumn();
    const hData = hachu.getRange(hBase, 1, hLastRow - hHeaderRow, hLastCol).getValues();

    // 発注データ側：商品コード＋業者ごとに、補完したい値を集める
    // 同じ商品コード＋業者が複数ある場合は、下の行の非空値を優先
    const sourceMap = new Map();

    for (let i = 0; i < hData.length; i++) {
      const row = hData[i];

      const code = String(row[hCodeCol - 1] || '').trim();
      const vendor = String(row[hVendorCol - 1] || '').trim();

      if (!code || !vendor) continue;

      const key = code + sep + vendor;
      const rec = sourceMap.get(key) || {};

      usableRules.forEach(rule => {
        const value = row[rule.sourceCol - 1];
        if (_補完_notBlank_(value)) {
          rec[rule.label] = value;
        }
      });

      sourceMap.set(key, rec);
    }

    const mRows = mLastRow - mHeaderRow;
    const maxMasterCol = Math.max(
      master.getLastColumn(),
      mCodeCol,
      mVendorCol,
      ...usableRules.map(r => r.masterCol)
    );

    const mData = master.getRange(mBase, 1, mRows, maxMasterCol).getValues();

    // 更新列ごとに配列を用意して、最後にまとめて書き戻す
    const colValuesMap = {};
    usableRules.forEach(rule => {
      colValuesMap[rule.masterCol] = master.getRange(mBase, rule.masterCol, mRows, 1).getValues();
    });

    let matched = 0;
    let updated = 0;
    let skipped = 0;
    let noSource = 0;

    const updateCountByLabel = {};

    for (let i = 0; i < mData.length; i++) {
      const mRow = mData[i];

      const code = String(mRow[mCodeCol - 1] || '').trim();
      const vendor = String(mRow[mVendorCol - 1] || '').trim();

      if (!code || !vendor) {
        skipped++;
        continue;
      }

      const key = code + sep + vendor;
      const src = sourceMap.get(key);

      if (!src) {
        noSource++;
        continue;
      }

      matched++;

      usableRules.forEach(rule => {
        const newValue = src[rule.label];

        if (!_補完_notBlank_(newValue)) return;

        const oldValue = mRow[rule.masterCol - 1];

        // overwrite=false の場合、マスタに既に値があれば触らない
        if (!rule.overwrite && _補完_notBlank_(oldValue)) {
          skipped++;
          return;
        }

        if (_補完_normCell_(oldValue) !== _補完_normCell_(newValue)) {
          colValuesMap[rule.masterCol][i][0] = newValue;
          mData[i][rule.masterCol - 1] = newValue;

          updated++;
          updateCountByLabel[rule.label] = (updateCountByLabel[rule.label] || 0) + 1;
        } else {
          skipped++;
        }
      });
    }

    // 一括書き戻し
    Object.keys(colValuesMap).forEach(col => {
      master.getRange(mBase, Number(col), mRows, 1).setValues(colValuesMap[col]);
    });

    let msg =
      '商品マスタの既存データ補完が完了したで\n\n' +
      `発注側キー数：${sourceMap.size}\n` +
      `マスタ一致行：${matched}\n` +
      `更新：${updated}\n` +
      `変更なし・スキップ：${skipped}\n` +
      `発注側に元データなし：${noSource}`;

    const labels = Object.keys(updateCountByLabel);
    if (labels.length) {
      msg += '\n\n【更新内訳】';
      labels.forEach(label => {
        msg += `\n${label}：${updateCountByLabel[label]}`;
      });
    }

    if (missingRules.length) {
      msg += '\n\n⚠ 見つからなかった補完ルール';
      missingRules.forEach(r => {
        msg += `\n${r.label}：発注列 or マスタ列が見つからんかった`;
      });
    }

    say(msg);

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 * 補助関数
 ************************************************************/

function _補完_headerRows_(mainHeaderRow) {
  const rows = [];
  if (mainHeaderRow - 1 >= 1) rows.push(mainHeaderRow - 1);
  rows.push(mainHeaderRow);
  return rows;
}

function _補完_resolveCol_(sheet, cfgKey, headerNames, headerRows) {
  const cfg = (typeof CFG !== 'undefined') ? CFG : {};

  if (cfgKey && cfg[cfgKey]) {
    return cfg[cfgKey];
  }

  return _補完_findColByHeaders_(sheet, headerNames, headerRows);
}

function _補完_findColByHeaders_(sheet, headerNames, headerRows) {
  const lastCol = sheet.getLastColumn();
  const targets = headerNames.map(_補完_normHeader_);

  for (let r = 0; r < headerRows.length; r++) {
    const rowNo = headerRows[r];
    if (rowNo < 1) continue;

    const headers = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];

    for (let c = 0; c < headers.length; c++) {
      const h = _補完_normHeader_(headers[c]);
      if (targets.includes(h)) {
        return c + 1;
      }
    }
  }

  return 0;
}

function _補完_normHeader_(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/\s/g, '')
    .replace(/[（）]/g, '')
    .replace(/[()]/g, '');
}

function _補完_notBlank_(v) {
  return v !== '' && v !== null && typeof v !== 'undefined';
}

function _補完_normCell_(v) {
  if (v instanceof Date) {
    return v.getTime();
  }
  return String(v ?? '').trim();
}


function _masterMap(){
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.MASTER_SHEET), last=sh.getLastRow();
  const info={}, byKey={}, byKeyW={}, byCode={};
  if(last<=CFG.MASTER_HEADER_ROW) return {info,byKey,byKeyW,byCode};
  const n=last-CFG.MASTER_HEADER_ROW, base=CFG.MASTER_HEADER_ROW+1;
  const v=sh.getRange(base,1,n,sh.getLastColumn()).getValues();
  for(const r of v){
    const c=String(r[CFG.M_CODE-1]).trim(); if(!c) continue;
    const nm=r[CFG.M_NAME-1], it=r[CFG.M_ITEM-1];
    const ve=String(r[CFG.M_VENDOR-1]).trim(), p=r[CFG.M_PRICE-1];
    const w=r[CFG.M_WEIGHT-1];                                  // ★重さ
    if(!info[c]) info[c]={name:nm,item:it};
    byKey[c+CFG.SEP+ve]=p;
    byKeyW[c+CFG.SEP+ve]=w;                                     // ★重さ用の対応表
    (byCode[c]=byCode[c]||[]).push({vendor:ve, price:p, weight:w});
  }
  return {info,byKey,byKeyW,byCode};
}

function autofillHachu(e){
  const sh=e.range.getSheet(); if(sh.getName()!==CFG.HACHU_SHEET) return;
  const row=e.range.getRow(); if(row<=CFG.HACHU_HEADER_ROW) return;
  const col=e.range.getColumn(); if(col!==CFG.HACHU_CODE && col!==CFG.HACHU_VENDOR) return;

  const code=String(sh.getRange(row,CFG.HACHU_CODE).getValue()).trim(); if(!code) return;
  let vendor=String(sh.getRange(row,CFG.HACHU_VENDOR).getValue()).trim();

  const {info,byKey,byKeyW,byCode}=_masterMap();

  if(info[code]){
    sh.getRange(row,CFG.HACHU_NAME).setValue(info[code].name);
    sh.getRange(row,CFG.HACHU_ITEM).setValue(info[code].item);
  }

  let price=null, weight=null;
  if(vendor && byKey[code+CFG.SEP+vendor]!==undefined){
    price=byKey[code+CFG.SEP+vendor];
    weight=byKeyW[code+CFG.SEP+vendor];                         // ★
  } else if(byCode[code] && byCode[code].length===1){
    if(!vendor){ vendor=byCode[code][0].vendor; sh.getRange(row,CFG.HACHU_VENDOR).setValue(vendor); }
    price=byCode[code][0].price;
    weight=byCode[code][0].weight;                              // ★
  }
  if(price!==null && price!=='') sh.getRange(row,CFG.HACHU_PRICE).setValue(price);
  if(weight!==null && weight!=='' && typeof weight!=='undefined'){
    sh.getRange(row,CFG.HACHU_WEIGHT).setValue(weight);         // ★P列に重さ
  }
}

function registerToMaster(){
  const ss=SpreadsheetApp.getActive(), hachu=ss.getSheetByName(CFG.HACHU_SHEET), ui=SpreadsheetApp.getUi();
  const last=hachu.getLastRow(); if(last<=CFG.HACHU_HEADER_ROW){ui.alert('発注データが無いで');return;}
  const n=last-CFG.HACHU_HEADER_ROW, c=x=>hachu.getRange(CFG.HACHU_HEADER_ROW+1,x,n,1).getValues().map(r=>r[0]);
  const codes=c(CFG.HACHU_CODE), vendors=c(CFG.HACHU_VENDOR), names=c(CFG.HACHU_NAME), items=c(CFG.HACHU_ITEM), prices=c(CFG.HACHU_PRICE);
  const master=ss.getSheetByName(CFG.MASTER_SHEET);
  const {byKey}=_masterMap(); let add=0, up=0;

  for(let i=0;i<n;i++){
    const code=String(codes[i]).trim(); if(!code) continue;
    const vendor=String(vendors[i]).trim();
    const key=code+CFG.SEP+vendor;
    if(byKey[key]===undefined){
      const r=master.getLastRow()+1;
      master.getRange(r,CFG.M_CODE).setValue(code);
      master.getRange(r,CFG.M_NAME).setValue(names[i]);
      master.getRange(r,CFG.M_ITEM).setValue(items[i]);
      master.getRange(r,CFG.M_VENDOR).setValue(vendor);
      master.getRange(r,CFG.M_PRICE).setValue(prices[i]);
      if(CFG.M_NO) master.getRange(r,CFG.M_NO).setValue(r-CFG.MASTER_HEADER_ROW);
      byKey[key]=prices[i]; add++;
    } else if(CFG.UPDATE_PRICE_IF_DIFF && prices[i]!=='' && byKey[key]!==prices[i]){
      _updateMasterPrice(master,code,vendor,prices[i]); byKey[key]=prices[i]; up++;
    }
  }
  ui.alert(`登録完了\n商品マスタ 追加 ${add} / 価格更新 ${up}`);
}
function _updateMasterPrice(master,code,vendor,price){
  const last=master.getLastRow(), n=last-CFG.MASTER_HEADER_ROW; if(n<=0) return;
  const base=CFG.MASTER_HEADER_ROW+1;
  const cs=master.getRange(base,CFG.M_CODE,n,1).getValues();
  const vs=master.getRange(base,CFG.M_VENDOR,n,1).getValues();
  for(let i=0;i<n;i++) if(String(cs[i][0]).trim()===code && String(vs[i][0]).trim()===vendor){
    master.getRange(base+i,CFG.M_PRICE).setValue(price); return;
  }
}
function _nextHachuNo(sh, colNo){
  const start=CFG.HACHU_HEADER_ROW+1, last=sh.getLastRow();
  if(last<start) return 1;
  const vals=sh.getRange(start, colNo, last-start+1, 1).getValues();
  let max=0;
  vals.forEach(r=>{ const v=Number(r[0]); if(!isNaN(v)&&v>max) max=v; });
  return max+1;
}

function _masterNextRow_(master) {
  const start = CFG.MASTER_HEADER_ROW + 1;
  const last = Math.max(master.getLastRow(), start);

  const numRows = last - start + 1;
  const codes = master.getRange(start, CFG.M_CODE, numRows, 1).getValues();

  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0] || '').trim() === '') {
      return start + i;
    }
  }

  return last + 1;
}

// グループの基準列を選ぶ:3=発注日(C) / 7=購入No(G)
const HATCHU_GROUP_COL = 7;   // ← 購入Noで囲みたいなら 7 に変える

// 発注シート:基準列が同じ連続行を太枠で囲む(中の格子は残す)
function 発注_drawGroupBorders_(sh) {
  const startRow = HATCHU_CFG.START_ROW;          // 7
  const last = sh.getLastRow();
  if (last < startRow) return;
  const numRows = last - startRow + 1;
  const startCol = 1;
  const maxCol = Number(HATCHU_CFG.MAX_COL || HATCHU_CFG.BORDER_COLS || 28);

  const keys = sh.getRange(startRow, HATCHU_GROUP_COL, numRows, 1)
                 .getDisplayValues().map(r => String(r[0] || '').trim());

  let gStart = null, cur = '';
  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;
    const key = isEnd ? '' : keys[i];

    if (gStart === null) {
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      continue;
    }
    if (isEnd || key !== cur) {
      // グループの外枠だけ太線(縦線・横線=null で中の格子は触らへん)
      sh.getRange(startRow + gStart, startCol, i - gStart, maxCol)
        .setBorder(true, true, true, true, null, null,
                   '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      else { gStart = null; cur = ''; }
    }
  }
}

// ボタン用:格子を引き直してから太枠を付ける(まるごとリフレッシュ)
function 発注_グループ罫線() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== HATCHU_CFG.SHEET_NAME) {
    SpreadsheetApp.getUi().alert('発注シートで実行してな');
    return;
  }
  const numRows = sh.getLastRow() - HATCHU_CFG.START_ROW + 1;
  if (numRows > 0) HATCHU_processBordersAndCheckboxes_(sh, HATCHU_CFG.START_ROW, numRows);
  発注_drawGroupBorders_(sh);
  SpreadsheetApp.getActive().toast('グループの太枠を更新したで');
}

// onEditから呼ぶ用:監視列を触ったときだけ全体補完を起動
function 商品コードと業者と商品名と価格がそろったら自動でマスタ登録(e){
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.HACHU_SHEET) return;

  const sCol = e.range.getColumn(), eCol = e.range.getLastColumn();
  const watch = [CFG.HACHU_VENDOR, CFG.HACHU_NAME, CFG.HACHU_ITEM, CFG.HACHU_PRICE, CFG.HACHU_WEIGHT];
  const pastedFullRow = e.range.getNumColumns() > 1 &&
    [CFG.HACHU_CODE, CFG.HACHU_VENDOR, CFG.HACHU_NAME, CFG.HACHU_PRICE].some(w => sCol <= w && w <= eCol);

  if (!watch.some(w => sCol <= w && w <= eCol) && !pastedFullRow) return;

  マスタ自動補完_編集範囲_(e.range);
}

function マスタ自動補完_編集範囲_(range, silent) {
  if (!range) return 0;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return 0;

  try {
    const ss = SpreadsheetApp.getActive();
    const hachu = range.getSheet();
    const master = ss.getSheetByName(CFG.MASTER_SHEET);

    if (!hachu || !master || hachu.getName() !== CFG.HACHU_SHEET) return 0;

    const hStart = CFG.HACHU_HEADER_ROW + 1;
    const startRow = Math.max(hStart, range.getRow());
    const endRow = Math.min(hachu.getLastRow(), range.getLastRow());
    if (endRow < startRow) return 0;

    const width = Math.max(
      CFG.HACHU_CODE,
      CFG.HACHU_VENDOR,
      CFG.HACHU_NAME,
      CFG.HACHU_ITEM,
      CFG.HACHU_PRICE,
      CFG.HACHU_WEIGHT || 0
    );

    const rows = hachu.getRange(startRow, 1, endRow - startRow + 1, width).getValues();
    const candidates = [];

    for (const r of rows) {
      const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
      const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
      const name = String(r[CFG.HACHU_NAME - 1] || '').trim();
      const item = String(r[CFG.HACHU_ITEM - 1] || '').trim();
      const price = r[CFG.HACHU_PRICE - 1];
      const weight = CFG.HACHU_WEIGHT ? r[CFG.HACHU_WEIGHT - 1] : '';

      if (!code || !vendor || !name) continue;
      if (price === '' || price === null || typeof price === 'undefined') continue;

      candidates.push([code, vendor, name, item, price, weight]);
    }

    if (candidates.length === 0) return 0;

    const { seen } = _masterSeen_(master);
    const toAdd = [];

    for (const row of candidates) {
      const key = row[0] + CFG.SEP + row[1];
      if (seen.has(key)) continue;
      seen.add(key);
      toAdd.push(row);
    }

    if (toAdd.length === 0) return 0;

    const writeRow = _masterNextRow_(master);
    master.getRange(writeRow, CFG.M_CODE, toAdd.length, 6).setValues(toAdd);

    if (!silent) {
      SpreadsheetApp.getActive().toast(`商品マスタに自動登録：${toAdd.length}件`);
    }

    return toAdd.length;
  } finally {
    lock.releaseLock();
  }
}

function マスタ自動補完_(silent) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return 0;

  try {
    const ss = SpreadsheetApp.getActive();
    const hachu  = ss.getSheetByName(CFG.HACHU_SHEET);
    const master = ss.getSheetByName(CFG.MASTER_SHEET);

    if (!hachu || !master) return 0;

    const { seen } = _masterSeen_(master);

    const hStart = CFG.HACHU_HEADER_ROW + 1;
    const hLast  = hachu.getLastRow();
    if (hLast < hStart) return 0;

    SpreadsheetApp.flush();

    const hv = hachu
      .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
      .getValues();

    const toAdd = []; // [code, vendor, name, item, price, weight]

    for (const r of hv) {
      const code   = String(r[CFG.HACHU_CODE   - 1] || '').trim();
      const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
      const name   = String(r[CFG.HACHU_NAME   - 1] || '').trim();
      const item   = String(r[CFG.HACHU_ITEM   - 1] || '').trim();
      const price  = r[CFG.HACHU_PRICE  - 1];
      const weight = CFG.HACHU_WEIGHT ? r[CFG.HACHU_WEIGHT - 1] : '';

      if (!code || !vendor || !name) continue;
      if (price === '' || price === null) continue;

      const key = code + CFG.SEP + vendor;
      if (seen.has(key)) continue;

      seen.add(key);
      toAdd.push([code, vendor, name, item, price, weight]);
    }

    if (toAdd.length === 0) return 0;

    const writeRow = _masterNextRow_(master);

    // B〜Gへ登録
    master.getRange(writeRow, CFG.M_CODE, toAdd.length, 6)
      .setValues(toAdd);

    if (!silent) {
      SpreadsheetApp.getActive().toast(`商品マスタに自動登録：${toAdd.length}件`);
    }

    return toAdd.length;

  } finally {
    lock.releaseLock();
  }
}

function 商品マスタ_発注シートから今すぐ補完登録() {
  const n = マスタ自動補完_(false);
  SpreadsheetApp.getUi().alert(`商品マスタ補完完了：${n}件追加`);
}

function _masterSeen_(master) {
  const seen = new Set();

  const base = CFG.MASTER_HEADER_ROW + 1;
  const last = master.getLastRow();

  if (last < base) {
    return { seen };
  }

  const n = last - CFG.MASTER_HEADER_ROW;
  const vals = master.getRange(base, 1, n, master.getLastColumn()).getValues();

  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][CFG.M_CODE - 1] || '').trim();
    const vendor = String(vals[i][CFG.M_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    seen.add(code + CFG.SEP + vendor);
  }

  return { seen };
}

function 商品マスタ_重さだけ発注から補完更新() {
  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!hachu || !master) {
    ui.alert('発注シートか商品マスタが見つからんで');
    return;
  }

  const hStart = CFG.HACHU_HEADER_ROW + 1;
  const hLast = hachu.getLastRow();
  if (hLast < hStart) {
    ui.alert('発注データが無いで');
    return;
  }

  const hVals = hachu
    .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
    .getValues();

  // 発注側から「商品コード＋業者 → 重さ」を作る
  const weightMap = new Map();

  for (const r of hVals) {
    const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
    const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
    const weight = r[CFG.HACHU_WEIGHT - 1];

    if (!code || !vendor) continue;
    if (weight === '' || weight === null || typeof weight === 'undefined') continue;

    weightMap.set(code + CFG.SEP + vendor, weight);
  }

  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mLast = master.getLastRow();
  if (mLast < mStart) {
    ui.alert('商品マスタにデータが無いで');
    return;
  }

  const mRows = mLast - CFG.MASTER_HEADER_ROW;
  const mVals = master
    .getRange(mStart, 1, mRows, master.getLastColumn())
    .getValues();

  const weightValues = master
    .getRange(mStart, CFG.M_WEIGHT, mRows, 1)
    .getValues();

  let updated = 0;
  let matched = 0;

  for (let i = 0; i < mVals.length; i++) {
    const code = String(mVals[i][CFG.M_CODE - 1] || '').trim();
    const vendor = String(mVals[i][CFG.M_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;
    if (!weightMap.has(key)) continue;

    matched++;

    const newWeight = weightMap.get(key);
    const oldWeight = weightValues[i][0];

    if (String(oldWeight || '').trim() !== String(newWeight || '').trim()) {
      weightValues[i][0] = newWeight;
      updated++;
    }
  }

  master.getRange(mStart, CFG.M_WEIGHT, mRows, 1).setValues(weightValues);

  ui.alert(
    `商品マスタの重さ補完完了\n` +
    `発注側 重さデータ：${weightMap.size}件\n` +
    `商品マスタ一致：${matched}件\n` +
    `重さ更新：${updated}件`
  );
}

function 商品マスタ_足りないデータだけ発注から補完(silent) {
  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  // トリガー実行時はUIが使えないのでログに切り替える
  const say = msg => { if (silent) Logger.log(msg); else SpreadsheetApp.getUi().alert(msg); };

  if (!hachu || !master) {
    say('発注シートか商品マスタが見つからんで');
    return;
  }

  const hStart = CFG.HACHU_HEADER_ROW + 1;
  const hLast = hachu.getLastRow();
  if (hLast < hStart) {
    say('発注データが無いで');
    return;
  }

  const hVals = hachu
    .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
    .getValues();

  // 発注シート側の 商品コード＋業者 → 補完元データ
  const srcMap = new Map();

  for (const r of hVals) {
    const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
    const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;

    const old = srcMap.get(key) || {};

    srcMap.set(key, {
      name: old.name || r[CFG.HACHU_NAME - 1],
      item: old.item || r[CFG.HACHU_ITEM - 1],
      price: old.price || r[CFG.HACHU_PRICE - 1],
      weight: old.weight || r[CFG.HACHU_WEIGHT - 1]
    });
  }

  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mLast = master.getLastRow();
  if (mLast < mStart) {
    say('商品マスタにデータが無いで');
    return;
  }

  const mRows = mLast - CFG.MASTER_HEADER_ROW;
  const mLastCol = master.getLastColumn();

  const mVals = master
    .getRange(mStart, 1, mRows, mLastCol)
    .getValues();

  let updated = 0;
  let matched = 0;

  for (let i = 0; i < mVals.length; i++) {
    const code = String(mVals[i][CFG.M_CODE - 1] || '').trim();
    const vendor = String(mVals[i][CFG.M_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;
    const src = srcMap.get(key);

    if (!src) continue;

    matched++;

    // 商品名：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_NAME - 1]) && !isBlankForMasterFill_(src.name)) {
      master.getRange(mStart + i, CFG.M_NAME).setValue(src.name);
      updated++;
    }

    // 品目：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_ITEM - 1]) && !isBlankForMasterFill_(src.item)) {
      master.getRange(mStart + i, CFG.M_ITEM).setValue(src.item);
      updated++;
    }

    // 価格：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_PRICE - 1]) && !isBlankForMasterFill_(src.price)) {
      master.getRange(mStart + i, CFG.M_PRICE).setValue(src.price);
      updated++;
    }

    // 重さ：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_WEIGHT - 1]) && !isBlankForMasterFill_(src.weight)) {
      master.getRange(mStart + i, CFG.M_WEIGHT).setValue(src.weight);
      updated++;
    }
  }

  say(
    '商品マスタの足りないデータ補完が完了したで\n\n' +
    `発注側データ：${srcMap.size}件\n` +
    `商品マスタ一致：${matched}件\n` +
    `補完したセル：${updated}件`
  );
}

function isBlankForMasterFill_(v) {
  return v === null ||
         typeof v === 'undefined' ||
         String(v).trim() === '' ||
         String(v).trim() === '　';
}

function テスト_STD093をマスタで探す() {
  const m = _masterMap();
  const code = 'STD093';
  const msg =
    'STD093 マスタ照合結果\n\n' +
    'info(名前/品目)：' + JSON.stringify(m.info[code]) + '\n' +
    'byCode(業者/価格/重さ)：' + JSON.stringify(m.byCode[code]);
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}