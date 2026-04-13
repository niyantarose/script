/******************************************************
 * 一気通貫：①(C=商品コード, D=種類数) → ②(A:D)作成(表示名Dは保持) →
 *           ②の表示名Dを使って Yahoo商品登録シートのF列(options)へ反映
 * 追加で sub-code も書きたければ設定でON
 *
 * ✅ 種類数=1：sub-code は codea/codeb（-1は付けない）
 * ✅ 種類数>=2：sub-code は code-1a/code-1b ...、options は \n\n で2行構成
 * ✅ 台湾判定：Yahoo商品登録シートの name（または商品名）列に「台湾/台灣」が含まれる場合
 *
 * ★ v3: ②の罫線が「前回の分が残る」問題を完全除去
 *      - ②は毎回「値クリア(A:D)」+「罫線クリア(全列)」→必要箇所だけ再描画
 ******************************************************/

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================
function IKKITUKAN_uiSafeAlert_(title, msg, hasOkButton) {
  try {
    var ui = SpreadsheetApp.getUi();
    if (hasOkButton) ui.alert(title, msg, ui.ButtonSet.OK);
    else ui.alert(String(title || msg));
    return;
  } catch (e) {}

  try {
    var fullMsg = title ? (title + '\n' + msg) : msg;
    SpreadsheetApp.getActiveSpreadsheet().toast(String(fullMsg), '一気通貫', 15);
  } catch (e2) {}

  Logger.log('[IKKITUKAN_UI_FALLBACK] ' + (title || '') + ' ' + (msg || ''));
}

function 一気通貫_1から2を作成してYahooF反映_台湾対応() {
  const CFG = {
    // シート名
    SHEET_1: '①商品入力シート',
    SHEET_2: '②オプション入力シート',
    SHEET_Y: 'Yahoo商品登録シート',

    // ①の列（固定運用：C=3, D=4）
    ROW_HEADER_1: 2,
    COL_CODE_1: 3,
    COL_KINDS_1: 4,

    // ②の構造
    ROW_HEADER_2: 1,      // ②は1行目ヘッダー
    OUT_COLS_2: 4,        // A:D

    // ★ ②の罫線クリア範囲（ここが“残る問題”の本丸）
    // true: ②のデータ部(2行目以降)の「全列」を罫線クリア（前回がA:D以外でも確実に消す）
    // false: A:Dだけ罫線クリア（軽いけど、前回が全列罫線だと残る）
    CLEAR_BORDERS_ALL_COLUMNS_ON_SHEET2: true,

    // 在庫ラベル
    SOKUNO: '即納（日本在庫）',
    KOREA:  'お取り寄せ（韓国から）',
    TAIWAN: 'お取り寄せ（台湾から）',

    // Yahoo：options / sub-code
    FALLBACK_OPTIONS_COL_Y: 6, // F列=6
    WRITE_SUBCODE_TOO: true,
    FALLBACK_SUBCODE_COL_Y: 4, // D列=4
  };

  const ss  = SpreadsheetApp.getActive();
  const sh1 = ss.getSheetByName(CFG.SHEET_1);
  const sh2 = ss.getSheetByName(CFG.SHEET_2);
  const shy = ss.getSheetByName(CFG.SHEET_Y);
  if (!sh1 || !sh2 || !shy) throw new Error('①/②/Yahoo商品登録シート のいずれかが見つかりません');

  // -----------------------------
  // ① 読み取り：code と kinds
  // -----------------------------
  const last1 = sh1.getLastRow();
  if (last1 <= CFG.ROW_HEADER_1) {
    IKKITUKAN_uiSafeAlert_('①にデータがありません');
    return;
  }

  const num1  = last1 - CFG.ROW_HEADER_1;
  const rows1 = sh1.getRange(CFG.ROW_HEADER_1 + 1, 1, num1, Math.max(CFG.COL_KINDS_1, CFG.COL_CODE_1)).getValues();

  const codeInfo = new Map();
  rows1.forEach(r => {
    const code = String(r[CFG.COL_CODE_1 - 1] ?? '').trim();
    if (!code || code === '商品コード' || code.toLowerCase() === 'code') return;

    let kinds = Number(String(r[CFG.COL_KINDS_1 - 1] ?? '').trim());
    if (!isFinite(kinds) || kinds <= 0) kinds = 1;

    if (!codeInfo.has(code)) codeInfo.set(code, { kinds });
    else codeInfo.get(code).kinds = Math.max(codeInfo.get(code).kinds, kinds);
  });

  if (codeInfo.size === 0) {
    IKKITUKAN_uiSafeAlert_('①のC列に商品コードが見つかりません');
    return;
  }

  // -----------------------------------------
  // ② 既存の表示名(D)を保持するために先読み
  // -----------------------------------------
  const keepNameMap = readDisplayNameMapFrom_(sh2, CFG);

  // -----------------------------------------
  // ② を作成（A:D）※D列は保持
  // -----------------------------------------
  const out2 = [];
  const codesSorted = Array.from(codeInfo.keys()).sort((a, b) => a.localeCompare(b, 'ja'));

  let prev = null;
  codesSorted.forEach(code => {
    const kinds = codeInfo.get(code).kinds;

    if (prev !== null && prev !== code) out2.push(['', '', '', '']); // 見やすい空行
    prev = code;

    if (kinds === 1) {
      const name1 = keepNameMap.get(`${code}::1`) || '';
      out2.push([code, CFG.SOKUNO, `${code}a`, name1]);
      out2.push([code, CFG.KOREA,  `${code}b`, name1]); // 表示だけ。台湾はYahoo出力時に差し替える
      return;
    }

    for (let k = 1; k <= kinds; k++) {
      const keep = keepNameMap.get(`${code}::${k}`) || '';
      out2.push([code, CFG.SOKUNO, `${code}-${k}a`, keep]);
      out2.push([code, CFG.KOREA,  `${code}-${k}b`, keep]);
    }
  });

  // =====================================================
  // ★②クリア：値(A:D) + 罫線(全列) を「maxRowsまで」確実に消す
  // =====================================================
  IKKITUKAN_clearSheet2DataAndBorders_(sh2, CFG);

  // ②へ書き込み
  const startRow2 = CFG.ROW_HEADER_2 + 1;
  if (out2.length > 0) {
    sh2.getRange(startRow2, 1, out2.length, CFG.OUT_COLS_2).setValues(out2);
  }

  // ★「お取り寄せ」行の下に太罫線を付ける（必要箇所だけ再描画）
  _お取り寄せ行の下に太罫線_自動(out2, sh2, CFG);

  // ②の最新状態から表示名を読み直す
  const optIndex = buildOptionIndexFrom_(sh2, CFG);

  // -----------------------------------------
  // Yahoo 側の列位置を把握（ヘッダー検出＋フォールバック）
  // -----------------------------------------
  const yData = shy.getDataRange().getValues();
  if (yData.length < 2) throw new Error('Yahoo商品登録シートにデータがありません');

  const headerRowIndex = detectHeaderRowIndex_(yData);
  const header = (yData[headerRowIndex] || []).map(x => String(x || '').trim().toLowerCase());

  const idxCode = header.indexOf('code');
  const idxName = header.indexOf('name') !== -1 ? header.indexOf('name') : header.indexOf('商品名');
  const idxOpt  = header.indexOf('options') !== -1 ? header.indexOf('options') : header.indexOf('option');
  const idxSub  =
    header.indexOf('sub-code') !== -1 ? header.indexOf('sub-code') :
    header.indexOf('subcode')  !== -1 ? header.indexOf('subcode')  : -1;

  if (idxCode === -1) throw new Error('Yahoo商品登録シートに code 列が見つかりません');

  const colOptionsY = (idxOpt !== -1) ? (idxOpt + 1) : CFG.FALLBACK_OPTIONS_COL_Y;
  const colSubcodeY = (idxSub !== -1) ? (idxSub + 1) : CFG.FALLBACK_SUBCODE_COL_Y;

  const startRowY = headerRowIndex + 2; // 1-based
  const lastY = shy.getLastRow();
  const numY  = lastY - (startRowY - 1);
  if (numY <= 0) return;

  const yCodes = shy.getRange(startRowY, idxCode + 1, numY, 1).getValues();
  const yNames = (idxName !== -1)
    ? shy.getRange(startRowY, idxName + 1, numY, 1).getValues()
    : Array.from({ length: numY }, () => ['']);

  const outOptions = Array.from({ length: numY }, () => ['']);
  const outSubcode = Array.from({ length: numY }, () => ['']);

  let updated = 0;
  const missingNames = [];

  for (let i = 0; i < numY; i++) {
    const code = String(yCodes[i][0] || '').trim();
    if (!code || !codeInfo.has(code)) continue;

    const kinds = codeInfo.get(code).kinds;
    const productName = String(yNames[i][0] || '');
    const isTaiwan = /台湾|台灣/.test(productName);
    const labelB = isTaiwan ? CFG.TAIWAN : CFG.KOREA;

    const namesByKind = [];
    for (let k = 1; k <= kinds; k++) {
      const nm = optIndex.nameMap.get(`${code}::${k}`) || '';
      namesByKind.push(nm);
      if (kinds >= 2 && !nm) missingNames.push(`${code} / 種類${k}`);
    }

    let optionsStr = `★在庫の設定 ${CFG.SOKUNO} ${labelB}`;
    if (kinds >= 2) {
      const list = namesByKind.map((n, idx) => `${idx + 1}.${n || ('種類' + (idx + 1))}`).join(' ');
      optionsStr += `\n\n★種類の選択 ${list}`;
    }
    outOptions[i][0] = optionsStr;

    if (CFG.WRITE_SUBCODE_TOO) {
      let subStr = '';
      if (kinds === 1) {
        subStr =
          `★在庫の設定:${CFG.SOKUNO}=${code}a` +
          `&★在庫の設定:${labelB}=${code}b`;
      } else {
        const aList = [];
        const bList = [];
        for (let k = 1; k <= kinds; k++) {
          const disp = namesByKind[k - 1] || `種類${k}`;
          const subA = `${code}-${k}a`;
          const subB = `${code}-${k}b`;
          aList.push(`★在庫の設定:${CFG.SOKUNO}#★種類の選択:${k}.${disp}=${subA}`);
          bList.push(`★在庫の設定:${labelB}#★種類の選択:${k}.${disp}=${subB}`);
        }
        subStr = [...aList, ...bList].join('&');
      }
      outSubcode[i][0] = subStr;
    }

    updated++;
  }

  shy.getRange(startRowY, colOptionsY, numY, 1).setValues(outOptions);
  if (CFG.WRITE_SUBCODE_TOO) {
    shy.getRange(startRowY, colSubcodeY, numY, 1).setValues(outSubcode);
  }

  if (missingNames.length > 0) {
    IKKITUKAN_uiSafeAlert_(
      '完了（ただし要確認）',
      `Yahooへ反映はしました。\n\n` +
      `更新件数: ${updated}\n` +
      `※種類が2以上なのに②のD列（表示名）が空の箇所があります。\n` +
      `空欄は「種類1/2…」で出力しています。\n\n` +
      `例:\n- ${missingNames.slice(0, 10).join('\n- ')}${missingNames.length > 10 ? '\n…' : ''}`,
      true
    );
  } else {
    IKKITUKAN_uiSafeAlert_('完了', `②を書き出し → Yahoo(F列)へ反映しました。\n更新件数: ${updated}`, true);
  }
}

/**
 * ★②シート：データ部(2行目以降)の「値(A:D)」と「罫線(全列 or A:D)」を maxRows まで確実に消す
 * これで“前回の太罫線が残る”を根絶できる
 */
function IKKITUKAN_clearSheet2DataAndBorders_(sh2, CFG) {
  const startRow = CFG.ROW_HEADER_2 + 1;
  const rows     = sh2.getMaxRows() - CFG.ROW_HEADER_2;
  if (rows <= 0) return;

  const colsForContent = CFG.OUT_COLS_2; // A:D
  const colsForBorders = CFG.CLEAR_BORDERS_ALL_COLUMNS_ON_SHEET2 ? sh2.getMaxColumns() : CFG.OUT_COLS_2;

  // でかいシートでも落ちにくいように分割
  const CHUNK = 1000;
  for (let offset = 0; offset < rows; offset += CHUNK) {
    const n = Math.min(CHUNK, rows - offset);

    // 値はA:Dだけ
    sh2.getRange(startRow + offset, 1, n, colsForContent).clearContent();

    // 罫線は指定列幅（基本：全列）
    sh2.getRange(startRow + offset, 1, n, colsForBorders)
       .setBorder(false, false, false, false, false, false);
  }
}

/** ②から「code::kindNo -> 表示名」を読む（既存Dの保持用） */
function readDisplayNameMapFrom_(sh2, CFG) {
  const map = new Map();
  const last = sh2.getLastRow();
  if (last <= CFG.ROW_HEADER_2) return map;

  const values = sh2.getRange(CFG.ROW_HEADER_2 + 1, 1, last - CFG.ROW_HEADER_2, 4).getValues();
  values.forEach(r => {
    const code = String(r[0] || '').trim();
    const sub  = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();
    if (!code || !sub) return;

    const kind = kindNoFromSub_(code, sub);
    map.set(`${code}::${kind}`, name);
  });
  return map;
}

/** ②の最新状態から index を作る（code別＆種類別の表示名取得用） */
function buildOptionIndexFrom_(sh2, CFG) {
  const res = { nameMap: new Map() };
  const last = sh2.getLastRow();
  if (last <= CFG.ROW_HEADER_2) return res;

  const values = sh2.getRange(CFG.ROW_HEADER_2 + 1, 1, last - CFG.ROW_HEADER_2, 4).getValues();
  values.forEach(r => {
    const code = String(r[0] || '').trim();
    const sub  = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();
    if (!code || !sub) return;

    const kind = kindNoFromSub_(code, sub);
    const key = `${code}::${kind}`;
    if (!res.nameMap.has(key) || (res.nameMap.get(key) === '' && name)) res.nameMap.set(key, name);
  });

  return res;
}

/** subCode から種類番号を推定 */
function kindNoFromSub_(code, sub) {
  const esc = String(code).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp('^' + esc + '-(\\d+)[ab]$', 'i');
  const m = String(sub).match(re);
  if (m) return Number(m[1]);

  // codea / codeb の場合は種類1
  const re1 = new RegExp('^' + esc + '[ab]$', 'i');
  if (re1.test(String(sub))) return 1;

  return 1;
}

/** Yahooヘッダー行が1行目か2行目か推定 */
function detectHeaderRowIndex_(yData) {
  const h1 = (yData[0] || []).map(x => String(x || '').trim().toLowerCase());
  const h2 = (yData[1] || []).map(x => String(x || '').trim().toLowerCase());
  const hasCode1 = h1.includes('code');
  const hasCode2 = h2.includes('code');
  if (hasCode2 && !hasCode1) return 1;
  return 0;
}

/**
 * ②のデータ部（A:D）に対して「お取り寄せ」行の下に太めの下罫線を引く
 */
function _お取り寄せ行の下に太罫線_自動(out2, sh2, CFG) {
  if (!out2 || out2.length === 0) return;

  const startRow = CFG.ROW_HEADER_2 + 1;
  const a1s = [];

  for (let i = 0; i < out2.length; i++) {
    const stockType = String(out2[i][1] || '');
    if (stockType && /お取り寄せ/.test(stockType)) {
      const r = startRow + i;
      a1s.push(`A${r}:D${r}`);
    }
  }
  if (a1s.length === 0) return;

  sh2.getRangeList(a1s).setBorder(
    null, null, true, null, null, null,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}
