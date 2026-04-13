/******************************************************
 * ①商品入力シート + ②オプション入力シート をもとに
 * Yahoo商品登録シート へ書き出し → その後 Yahoo用CSV成形（出力シートのみ）
 *
 * ・①商品入力シート：ヘッダー2行 → 3行目以降がデータ
 * ・②オプション入力シート：ヘッダー1行 → 2行目以降がデータ（A:D）
 * ・Yahoo商品登録シート：ヘッダー1行 → 2行目以降がデータ
 *
 * ★重要（今回の修正点）
 * - release-date（L列）は Yahoo仕様「半角数字のみ・YYYYMMDD」固定
 * - Google Sheetsの「テーブル（型付き列）」対応：
 *   -> setNumberFormat() を一切使わない（例外になるため）
 * - options(F列) 内の改行は \n のみ（\n\n は作らない & 後段でも潰す）
 * - options(F列) と sub-code(D列) の “項目名” は完全一致
 ******************************************************/

/** ======================================================
 * 設定
 * ====================================================== */
const YCFG = {
  SHEET_1: '①商品入力シート',
  SHEET_2: '②オプション入力シート',
  SHEET_Y: 'Yahoo商品登録シート',

  INPUT_HEADER_ROWS: 2,   // ①
  OPT_HEADER_ROWS: 1,     // ②
  OUT_HEADER_ROWS: 1,     // Yahoo商品登録シート

  // ①の列（A=1）
  COL_PATH: 1,      // A
  COL_NAME: 2,      // B
  COL_CODE: 3,      // C
  COL_PRICE: 5,     // E
  COL_HEADLINE: 7,  // G
  COL_CAPTION: 8,   // H
  COL_ABSTRACT: 9,  // I
  COL_EXPLAIN: 10,  // J
  COL_WEIGHT: 11,   // K
  COL_RELEASE: 12,  // L   ← release-date
  COL_TEMPLATE: 13, // M
  COL_PRODCAT: 14,  // N
  COL_JAN: 15,      // O
  COL_DISPLAY: 16,  // P
  COL_SPADD: 17,    // Q
  COL_POSTAGE: 18,  // R

  // 出力（A〜R=18列）だけ触る（S列以降は残す）
  WRITE_COLS: 18,

  // Yahooの項目名（options/sub-codeで必ず一致させる）
  OPT_STOCK: '★在庫の設定',
  OPT_KIND:  '★種類の選択',

  // 在庫ラベル（②が空の時の簡易パターン用）
  LABEL_SOKUNO: '即納（日本在庫）',
  LABEL_KOREA:  'お取り寄せ（韓国から）',
  LABEL_TAIWAN: 'お取り寄せ（台湾から）',
};


/** ======================================================
 * メイン：①→Yahoo商品登録シートへ書き出し → CSV成形
 * ====================================================== */
function 商品入力からYahooに反映してCSV成形_一括() {
  const ss = SpreadsheetApp.getActive();

  const 入力 = ss.getSheetByName(YCFG.SHEET_1);
  const opt  = ss.getSheetByName(YCFG.SHEET_2);
  const 出力 = ss.getSheetByName(YCFG.SHEET_Y);

  if (!入力) throw new Error('①商品入力シート が見つかりません。');
  if (!opt)  throw new Error('②オプション入力シート が見つかりません。');
  if (!出力) throw new Error('Yahoo商品登録シート が見つかりません。');

  const lastRow = 入力.getLastRow();

  // 入力にデータが無ければ：出力データ部(A〜Rのみ)をクリアして終了（S列以降は残す）
  if (lastRow <= YCFG.INPUT_HEADER_ROWS) {
    const outLast = 出力.getLastRow();
    if (outLast > YCFG.OUT_HEADER_ROWS) {
      出力.getRange(YCFG.OUT_HEADER_ROWS + 1, 1, outLast - YCFG.OUT_HEADER_ROWS, YCFG.WRITE_COLS).clearContent();
    }
    ss.toast('入力データなし：出力(A〜R)をクリアしました', '完了', 5);
    return;
  }

  const dataRowCount = lastRow - YCFG.INPUT_HEADER_ROWS;

  // ★ 入力はA〜Rだけ取れば十分（画像は別ルート）
  const srcValues = 入力.getRange(YCFG.INPUT_HEADER_ROWS + 1, 1, dataRowCount, YCFG.WRITE_COLS).getValues();

  // ②オプションを先に全部Map化（高速 + ぶれない）
  const optMap = buildOptMapFromSheet2_(opt);

  const outRows = [];

  for (let i = 0; i < dataRowCount; i++) {
    const r = srcValues[i];

    const code = String(r[YCFG.COL_CODE - 1] || '').trim();
    if (!code) continue;

    const path         = r[YCFG.COL_PATH - 1];
    const name         = r[YCFG.COL_NAME - 1];
    const price        = r[YCFG.COL_PRICE - 1];
    const headline     = r[YCFG.COL_HEADLINE - 1];
    const caption      = r[YCFG.COL_CAPTION - 1];
    const abstractTxt  = r[YCFG.COL_ABSTRACT - 1];
    const explanation  = r[YCFG.COL_EXPLAIN - 1];
    const shipWeight   = r[YCFG.COL_WEIGHT - 1];
    const releaseRaw   = r[YCFG.COL_RELEASE - 1];
    const template     = r[YCFG.COL_TEMPLATE - 1];
    const productCat   = r[YCFG.COL_PRODCAT - 1];
    const jan          = r[YCFG.COL_JAN - 1];
    const display      = r[YCFG.COL_DISPLAY - 1];
    const spAdditional = r[YCFG.COL_SPADD - 1];
    const postageSet   = r[YCFG.COL_POSTAGE - 1];

    // ★ 台湾判定：①商品入力シートB列（name）に「台湾/台灣」を含むか
    const isTaiwan = /台湾|台灣/.test(String(name || ''));
    const otoriLabel = isTaiwan ? YCFG.LABEL_TAIWAN : YCFG.LABEL_KOREA;

    // ✅ release-date を最初から Yahoo仕様(YYYYMMDD)の「文字列」に確定
    const releaseDate = normalizeDateYahoo8_(releaseRaw);

    // ②から options/sub-code を生成（完全一致ルール）
    let optionsHeader = ''; // F列（options）
    let subCodeStr    = ''; // D列（sub-code）

    const rowsForCode = optMap.get(code) || [];

    if (rowsForCode.length === 0) {
      // ②が無い場合：簡易（a/bのみ）
      const subA = code + 'a';
      const subB = code + 'b';
      subCodeStr =
        `${YCFG.OPT_STOCK}:${YCFG.LABEL_SOKUNO}=${subA}` +
        `&${YCFG.OPT_STOCK}:${otoriLabel}=${subB}`;
      optionsHeader =
        `${YCFG.OPT_STOCK} ${YCFG.LABEL_SOKUNO} ${otoriLabel}`;
    } else {
      const built = buildYahooOptionsFromSheet2Rows_(rowsForCode, isTaiwan);
      optionsHeader = built.optionsHeader;
      subCodeStr    = built.subCodeStr;

      // 念のため保険：韓国→台湾差し替え（②に韓国表記が残ってても通す）
      if (isTaiwan) {
        optionsHeader = optionsHeader.replace(/お取り寄せ（韓国から）/g, 'お取り寄せ（台湾から）');
        subCodeStr    = subCodeStr.replace(/お取り寄せ（韓国から）/g, 'お取り寄せ（台湾から）');
      }
    }

    // ★ A〜R（18列）だけ作る
    outRows.push([
      path,          // A
      name,          // B
      code,          // C
      subCodeStr,    // D   ← sub-code
      price,         // E
      optionsHeader, // F   ← options
      headline,      // G
      caption,       // H
      abstractTxt,   // I
      explanation,   // J
      shipWeight,    // K
      releaseDate,   // L   ← release-date（YYYYMMDD文字列）
      template,      // M
      productCat,    // N
      jan,           // O
      display,       // P
      spAdditional,  // Q
      postageSet     // R
    ]);
  }

  // 出力クリア(A〜Rのみ)して書き込み（S列以降は残す）
  const outLast = 出力.getLastRow();
  if (outLast > YCFG.OUT_HEADER_ROWS) {
    出力.getRange(YCFG.OUT_HEADER_ROWS + 1, 1, outLast - YCFG.OUT_HEADER_ROWS, YCFG.WRITE_COLS).clearContent();
  }

  if (outRows.length === 0) {
    ss.toast('書き出す行がありません（code空の行のみ）', '完了', 6);
    return;
  }

  出力.getRange(YCFG.OUT_HEADER_ROWS + 1, 1, outRows.length, YCFG.WRITE_COLS).setValues(outRows);

  // CSV成形（A〜Rの範囲内の列だけ触る）
  const rowsFormatted = formatOneSheet_(出力, YCFG.OUT_HEADER_ROWS);

  ss.toast(
    '反映→CSV成形 完了\n' +
    '書き出し：' + outRows.length + '行\n' +
    '成形対象：' + rowsFormatted + '行（Yahoo商品登録シート）',
    '完了',
    7
  );
}


/** ======================================================
 * ②オプション入力シートを code -> rows にする
 * rows: [{stockLabel, sub, name, kindNo, suffix}]
 * ====================================================== */
function buildOptMapFromSheet2_(optSheet) {
  const map = new Map();

  const last = optSheet.getLastRow();
  if (last <= YCFG.OPT_HEADER_ROWS) return map;

  const values = optSheet.getRange(YCFG.OPT_HEADER_ROWS + 1, 1, last - YCFG.OPT_HEADER_ROWS, 4).getValues();
  for (const r of values) {
    const code = String(r[0] || '').trim();
    const stockLabel = String(r[1] || '').trim();
    const sub = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();

    if (!code || !sub || !stockLabel) continue;

    const kindNo = kindNoFromSub_(code, sub); // 1..N
    const suffix = suffixFromSub_(code, sub); // a/b/other

    if (!map.has(code)) map.set(code, []);
    map.get(code).push({ stockLabel, sub, name, kindNo, suffix });
  }

  return map;
}

function kindNoFromSub_(code, sub) {
  const esc = String(code).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // code-3a / code-3b
  let m = String(sub).match(new RegExp('^' + esc + '-(\\d+)[ab]$', 'i'));
  if (m) return Number(m[1]);

  // codea / codeb
  if (new RegExp('^' + esc + '[ab]$', 'i').test(String(sub))) return 1;

  // それ以外は1扱い
  return 1;
}

function suffixFromSub_(code, sub) {
  const s = String(sub);
  if (/[aA]$/.test(s)) return 'a';
  if (/[bB]$/.test(s)) return 'b';
  return '';
}


/** ======================================================
 * ②の rows から Yahoo用 options(F) と sub-code(D) を生成
 * ★ここが「整合」の核
 * ====================================================== */
function buildYahooOptionsFromSheet2Rows_(rows, isTaiwan) {
  // stockLabel の順序：まず即納、それ以外
  const stockOrder = (label) => {
    if (label.includes('即納')) return 0;
    if (label.includes('お取り寄せ')) return 1;
    return 9;
  };

  // label 正規化
  const fixLabel = (label) => {
    let t = String(label || '');
    if (isTaiwan) t = t.replace(/お取り寄せ（韓国から）/g, 'お取り寄せ（台湾から）');
    return t;
  };

  // stockLabel のユニーク
  const labels = [];
  const labelSet = new Set();

  // kind のユニーク
  let maxKind = 1;
  const kindToName = new Map(); // kindNo -> 表示名

  for (const it of rows) {
    const lbl = fixLabel(it.stockLabel);
    if (!labelSet.has(lbl)) {
      labelSet.add(lbl);
      labels.push(lbl);
    }
    maxKind = Math.max(maxKind, Number(it.kindNo || 1));
    if (it.name) {
      if (!kindToName.has(it.kindNo) || kindToName.get(it.kindNo) === '') {
        kindToName.set(it.kindNo, it.name);
      }
    }
  }

  labels.sort((a, b) => stockOrder(a) - stockOrder(b));

  const kinds = maxKind;
  const hasKinds = kinds >= 2;

  // ---------- options(F列) ----------
  let optionsHeader = `${YCFG.OPT_STOCK} ` + labels.join(' ');

  if (hasKinds) {
    const parts = [];
    for (let k = 1; k <= kinds; k++) {
      const nm = kindToName.get(k) || `種類${k}`;
      parts.push(`${k}.${nm}`);
    }
    // ✅ \n\n は作らない：\n のみ
    optionsHeader += `\n${YCFG.OPT_KIND} ` + parts.join(' ');
  }

  // ---------- sub-code(D列) ----------
  let subCodeStr = '';

  if (!hasKinds) {
    // kinds=1 は簡易形式に落とす
    let subA = '';
    let subB = '';

    for (const it of rows) {
      const lbl = fixLabel(it.stockLabel);
      if (it.suffix === 'a' && (lbl.includes('即納') || stockOrder(lbl) === 0)) subA = it.sub;
      if (it.suffix === 'b' && lbl.includes('お取り寄せ')) subB = it.sub;
    }

    if (!subA) subA = (rows[0] && rows[0].sub) ? rows[0].sub : '';
    if (!subB) {
      const bRow = rows.find(x => String(x.stockLabel).includes('お取り寄せ') && x.suffix === 'b');
      if (bRow) subB = bRow.sub;
    }

    const labelA = labels.find(x => x.includes('即納')) || YCFG.LABEL_SOKUNO;
    const labelB = labels.find(x => x.includes('お取り寄せ')) || (isTaiwan ? YCFG.LABEL_TAIWAN : YCFG.LABEL_KOREA);

    subCodeStr =
      `${YCFG.OPT_STOCK}:${labelA}=${subA}` +
      `&${YCFG.OPT_STOCK}:${labelB}=${subB}`;

  } else {
    const items = [];

    for (const lbl0 of labels) {
      const lbl = fixLabel(lbl0);

      for (let k = 1; k <= kinds; k++) {
        const nm = kindToName.get(k) || `種類${k}`;

        const wantSuffix = lbl.includes('即納') ? 'a' : (lbl.includes('お取り寄せ') ? 'b' : '');
        let found = rows.find(x => fixLabel(x.stockLabel) === lbl && Number(x.kindNo) === k && (wantSuffix ? x.suffix === wantSuffix : true));
        if (!found) found = rows.find(x => fixLabel(x.stockLabel) === lbl && Number(x.kindNo) === k);
        const sub = found ? found.sub : '';

        items.push(`${YCFG.OPT_STOCK}:${lbl}#${YCFG.OPT_KIND}:${k}.${nm}=${sub}`);
      }
    }

    subCodeStr = items.join('&');
  }

  // 最終清掃
  optionsHeader = String(optionsHeader || '').replace(/ +$/gm, '');
  subCodeStr    = String(subCodeStr || '').replace(/\n/g, '').replace(/ +$/gm, '');

  return { optionsHeader, subCodeStr };
}

/** ======================================================
 * Yahoo用CSV成形（1シート版）
 *  - setNumberFormat() は使わない（テーブル型付き列で例外になるため）
 * ====================================================== */
function formatOneSheet_(sheet, headerRows) {
  const lastRow = sheet.getLastRow();
  const startRow = headerRows + 1;
  if (lastRow < startRow) return 0;

  const numRows = lastRow - headerRows;

  // 列番号（A=1）
  const COL_PRICE        = 5;   // E
  const COL_OPTIONS      = 6;   // F
  const COL_RELEASE_DATE = 12;  // L  ← 8桁にする
  const COL_TEMPLATE     = 13;  // M
  const COL_PRODUCT_CAT  = 14;  // N
  const COL_DISPLAY      = 16;  // P
  const COL_POSTAGE_SET  = 18;  // R

  const priceRange      = sheet.getRange(startRow, COL_PRICE,        numRows, 1);
  const optionsRange    = sheet.getRange(startRow, COL_OPTIONS,      numRows, 1);
  const releaseRange    = sheet.getRange(startRow, COL_RELEASE_DATE, numRows, 1);
  const templateRange   = sheet.getRange(startRow, COL_TEMPLATE,     numRows, 1);
  const productCatRange = sheet.getRange(startRow, COL_PRODUCT_CAT,  numRows, 1);
  const displayRange    = sheet.getRange(startRow, COL_DISPLAY,      numRows, 1);
  const postageSetRange = sheet.getRange(startRow, COL_POSTAGE_SET,  numRows, 1);

  const priceValues      = priceRange.getValues();
  const optionsValues    = optionsRange.getValues();
  const releaseValues    = releaseRange.getValues();
  const templateValues   = templateRange.getValues();
  const productCatValues = productCatRange.getValues();
  const displayValues    = displayRange.getValues();
  const postageSetValues = postageSetRange.getValues();

  for (let i = 0; i < numRows; i++) {
    priceValues[i][0]      = normalizePrice_(priceValues[i][0]);
    optionsValues[i][0]    = normalizeOptions_(optionsValues[i][0]);
    releaseValues[i][0]    = normalizeDateYahoo8_(releaseValues[i][0]); // ✅ YYYYMMDD
    templateValues[i][0]   = normalizeTemplate_(templateValues[i][0]);
    productCatValues[i][0] = normalizeInteger_(productCatValues[i][0]);
    displayValues[i][0]    = normalizeDisplay_(displayValues[i][0]);
    postageSetValues[i][0] = normalizeInteger_(postageSetValues[i][0]);
  }

  priceRange.setValues(priceValues);
  optionsRange.setValues(optionsValues);
  releaseRange.setValues(releaseValues);
  templateRange.setValues(templateValues);
  productCatRange.setValues(productCatValues);
  displayRange.setValues(displayValues);
  postageSetRange.setValues(postageSetValues);

  return numRows;
}


/** ======================================================
 * 正規化ヘルパー
 * ====================================================== */
function toHalfWidth_(str) {
  if (!str) return '';
  return String(str)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/．/g, '.')
    .replace(/－/g, '-')
    .replace(/／/g, '/');
}

function normalizeTemplate_(str) {
  if (!str) return '';
  str = toHalfWidth_(String(str).trim());
  return str.replace(/[^0-9A-Za-z]/g, '');
}

function normalizeInteger_(str) {
  if (str === null || str === undefined || str === '') return '';
  str = toHalfWidth_(str);
  str = str.replace(/[,，]/g, '');
  return str.replace(/[^\d]/g, '');
}

function normalizePrice_(str) {
  if (str === null || str === undefined || str === '') return '';
  str = toHalfWidth_(str);
  str = str.replace(/[,，]/g, '');
  str = str.replace(/[^0-9.]/g, '');
  if (!str) return '';
  const parts = str.split('.');
  if (parts.length > 2) str = parts[0] + '.' + parts.slice(1).join('');
  const num = Number(str);
  if (isNaN(num)) return '';
  return num;
}

function normalizeDisplay_(value) {
  const str = toHalfWidth_(String(value || '').trim());
  if (str === '1') return 1;
  if (str === '0' || str === '') return 0;
  return 0;
}

function normalizeOptions_(value) {
  let str = String(value || '');
  if (!str) return '';
  str = str.replace(/　/g, ' ');
  str = str.replace(/\t/g, '');

  // ✅ \n\nは禁止：全部単改行へ
  str = str.replace(/\n{2,}/g, '\n');

  // 行末スペース除去
  str = str.replace(/ +$/gm, '');
  return str;
}

/**
 * ✅ release-date を Yahoo仕様：半角8桁 YYYYMMDD に統一（文字列）
 * - 入力が Date型 / "2026/02/10" / "2026-2-10" / "20260210" 等を全部吸収
 * - 日付妥当性チェックあり
 * - 365日ルール：現在日付から 0〜365日以外は空欄（必要ならOFFに）
 */
function normalizeDateYahoo8_(value) {
  if (!value) return '';

  let y, mo, da;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    y  = value.getFullYear();
    mo = value.getMonth() + 1;
    da = value.getDate();
  } else {
    let str = toHalfWidth_(String(value)).trim();
    if (!str) return '';
    str = str.replace(/\s+/g, '');

    // 8桁（YYYYMMDD）
    let m = str.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) {
      y  = Number(m[1]);
      mo = Number(m[2]);
      da = Number(m[3]);
    } else {
      // 区切りあり
      str = str.replace(/[.\/]/g, '-');
      m = str.match(/^(\d{4})[\-年]?(\d{1,2})[\-月]?(\d{1,2})/);
      if (!m) return '';
      y  = Number(m[1]);
      mo = Number(m[2]);
      da = Number(m[3]);
    }
  }

  const dt = new Date(y, mo - 1, da);
  if (isNaN(dt.getTime()) || dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== da) return '';

  // 365日ルール（Yahoo仕様）：不要ならこのブロックをコメントアウト
  const today = new Date();
  today.setHours(0,0,0,0);
  const t = new Date(dt.getTime());
  t.setHours(0,0,0,0);

  const diffDays = Math.floor((t.getTime() - today.getTime()) / 86400000);
  if (diffDays < 0 || diffDays > 365) return '';

  const mm = ('0' + mo).slice(-2);
  const dd = ('0' + da).slice(-2);
  return String(y) + mm + dd;
}


/** ======================================================
 * （オマケ）②のD列を a行→b行へコピーして揃える
 * ====================================================== */
function 表示名をaからbへコピー_同バリエーション(fillBlankOnly) {
  const ss = SpreadsheetApp.getActive();
  const sh2 = ss.getSheetByName(YCFG.SHEET_2);
  if (!sh2) throw new Error('②オプション入力シート が見つかりません。');

  const last = sh2.getLastRow();
  if (last <= YCFG.OPT_HEADER_ROWS) return;

  const startRow = YCFG.OPT_HEADER_ROWS + 1;
  const numRows = last - YCFG.OPT_HEADER_ROWS;

  const vals = sh2.getRange(startRow, 1, numRows, 4).getValues(); // A:D
  const aName = new Map(); // key=code::kind -> name(from a)

  // 1) a側の名前を集める
  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][0] || '').trim();
    const sub  = String(vals[i][2] || '').trim();
    const name = String(vals[i][3] || '').trim();
    if (!code || !sub) continue;

    const kind = kindNoFromSub_(code, sub);
    const suf  = suffixFromSub_(code, sub);
    if (suf === 'a' && name) {
      aName.set(`${code}::${kind}`, name);
    }
  }

  // 2) b側を埋める
  let changed = 0;
  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][0] || '').trim();
    const sub  = String(vals[i][2] || '').trim();
    const name = String(vals[i][3] || '').trim();
    if (!code || !sub) continue;

    const kind = kindNoFromSub_(code, sub);
    const suf  = suffixFromSub_(code, sub);
    if (suf !== 'b') continue;

    const src = aName.get(`${code}::${kind}`) || '';
    if (!src) continue;

    if (fillBlankOnly) {
      if (!name) {
        vals[i][3] = src;
        changed++;
      }
    } else {
      if (name !== src) {
        vals[i][3] = src;
        changed++;
      }
    }
  }

  if (changed > 0) {
    sh2.getRange(startRow, 1, numRows, 4).setValues(vals);
  }

  ss.toast(`②表示名コピー 完了：${changed}件`, '完了', 6);
}
