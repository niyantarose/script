/******************************************************
 * 在庫検索（完成版）
 * - 手動実行（UIあり）: 在庫検索を実行()
 * - 自動実行（UIなし）: onEdit(e) → 在庫検索トリガー_(e)
 * - 本体ロジック（UIなし）: 在庫検索実行_(ss, sheet) ※処理件数を return
 *
 * 【前提】
 * - シート「在庫検索」: A列=検索コード, B列=即納, C列=取寄（1行目ヘッダー、2行目〜）
 * - シート「Yahoo全在庫」: A=code, C=sub-code, D=quantity（1行目ヘッダー、2行目〜）
 ******************************************************/

/**********************
 * 例外サフィックス判定
 **********************/
const 例外サフィックスマップ = {
  // BEAUTY2005 系の特殊パターン
  'BEAUTY2005aA': '即納',
  'BEAUTY2005bB': '即納',
  'BEAUTY2005Ab': '取寄',
  'BEAUTY2005Ba': '取寄',
};

/**********************
 * 完全に無視したい sub-code
 * （ARNA2302 系の小文字コードなど）
 **********************/
const 無視サフィックスコード = {
  'arna2302a': true,
  'arna2302b': true,
  'arna2302bb': true,
};

/**
 * sub-code が 即納か取寄かを判定する共通関数
 * @param {string} code     親コード
 * @param {string} subCode  サブコード
 * @return {'即納'|'取寄'|null}  null のときは集計対象外
 */
/**
 * sub-code が 即納か取寄かを判定する共通関数（拡張版）
 * @param {string} code     親コード
 * @param {string} subCode  サブコード
 * @return {'即納'|'取寄'|null}  null のときは集計対象外
 */
function 判定在庫種別_(code, subCode) {
  subCode = (subCode || '').toString().trim();
  code    = (code || '').toString().trim();

  if (!code) return '即納';

  // 0) 無視対象
  if (subCode) {
    const lower = subCode.toLowerCase();
    if (無視サフィックスコード[lower]) return null;
  }

  // 1) 例外テーブル（完全一致）
  if (subCode && 例外サフィックスマップ[subCode]) {
    return 例外サフィックスマップ[subCode];
  }

  // 2) sub-code 空（親コードだけの行）は即納扱い
  if (!subCode) return '即納';

  // 3) 末尾「1 / 2」で判定（1=即納、2=取寄）←追加
  const lastChar = subCode.slice(-1);
  if (lastChar === '1') return '即納';
  if (lastChar === '2') return '取寄';

  // 4) 末尾「a / b」で判定（b=取寄、それ以外=即納）
  const last = lastChar.toLowerCase();
  if (last === 'b') return '取寄';
  if (last === 'a') return '即納';

  // 5) それ以外は従来通り「即納」寄りに倒す（安全側）
  return '即納';
}


/****************************************
 * 在庫検索を手動実行（メニュー/ボタン用：UIあり）
 ****************************************/
function 在庫検索を実行() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('在庫検索');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('「在庫検索」シートが見つかりません');
    return;
  }

  const processed = 在庫検索実行_(ss, sheet); // UIなし本体
  SpreadsheetApp.getUi().alert('検索完了\n' + processed + '件を処理しました');
}

/****************************************
 * onEdit トリガー（A列を貼り替えたら自動検索）
 * ※UI(getUi)は絶対に使わない
 ****************************************/
function onEdit(e) {
  if (!e || !e.source || !e.range) return;

  try {
    在庫検索トリガー_(e);
  } catch (err) {
    console.error('onEdit エラー:', err);
    // UIは使えないのでログだけ
    Logger.log('onEdit エラー: ' + (err && err.message ? err.message : err));
  }
}

/****************************************
 * トリガー本体（UIなし）
 ****************************************/
function 在庫検索トリガー_(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  if (!sheet || sheet.getName() !== '在庫検索') return;

  const col = e.range.getColumn();
  if (col !== 1) return; // A列以外は無視

  const processed = 在庫検索実行_(ss, sheet);
  ss.toast(`在庫検索更新: ${processed}件`, '在庫検索', 3); // toastはOK
}

/****************************************
 * 在庫検索の実行ロジック（親コード集計対応版：UIなし）
 * @return {number} 処理件数（検索対象コード数）
 ****************************************/
function 在庫検索実行_(ss, sheet) {
  const yahooSheet = ss.getSheetByName('Yahoo全在庫');
  if (!yahooSheet) throw new Error('「Yahoo全在庫」シートが見つかりません。');

  const yahooLastRow = yahooSheet.getLastRow();
  if (yahooLastRow < 2) return 0;

  const searchSheetLastRow = sheet.getLastRow();
  if (searchSheetLastRow < 2) return 0;

  // ===== 背景クリア（A〜C）=====
  const clearRange = sheet.getRange(2, 1, searchSheetLastRow - 1, 3);
  clearRange.setBackground(null);

  // ===== A列の条件付き書式だけ削除 =====
  const rules = sheet.getConditionalFormatRules();
  const newRules = [];
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    const ranges = r.getRanges();
    const first = ranges && ranges.length ? ranges[0] : null;

    // 「A列にかかってるルール」は落とす（それ以外は残す）
    if (first && first.getColumn && first.getColumn() === 1) {
      // 捨てる
    } else {
      newRules.push(r);
    }
  }
  sheet.setConditionalFormatRules(newRules);

  // ===== 検索コード一覧（A列）→ Map 化 =====
  const searchCodesRange = sheet.getRange(2, 1, searchSheetLastRow - 1, 1);
  const searchCodesValues = searchCodesRange.getValues();

  const searchMap = {}; // code -> index
  for (let i = 0; i < searchCodesValues.length; i++) {
    const code = (searchCodesValues[i][0] || '').toString().trim();
    if (code) searchMap[code] = i;
  }

  const searchKeys = Object.keys(searchMap);
  Logger.log('在庫検索開始: ' + searchKeys.length + 'コード');
  if (searchKeys.length === 0) return 0;

  // ===== Yahoo全在庫を取得（A:code, C:sub-code, D:quantity）=====
  const yahooData = yahooSheet.getRange(2, 1, yahooLastRow - 1, 4).getValues();

  // ===== 親コードが「サフィックス付き行を持つか」事前チェック =====
  const hasVariant = {};
  for (let i = 0; i < yahooData.length; i++) {
    const row = yahooData[i];
    const rowCode = (row[0] || '').toString().trim();
    const subCode = (row[2] || '').toString().trim();
    if (!rowCode) continue;
    if (subCode && subCode !== rowCode) {
      hasVariant[rowCode] = true;
    }
  }

  // ===== 結果配列 [即納, 取寄]（検索行数ぶん）=====
  const results = Array.from({ length: searchCodesValues.length }, () => [0, 0]);

  // ===== 本番集計 =====
  for (let i = 0; i < yahooData.length; i++) {
    const row = yahooData[i];
    const rowCode = (row[0] || '').toString().trim();
    const subCode = (row[2] || '').toString().trim();
    const qty = Number(row[3]) || 0;

    if (!rowCode || qty === 0) continue;

    const kind = 判定在庫種別_(rowCode, subCode);
    if (!kind) continue; // 無視対象

    // ① sub-code で検索している行（MAXM2511Aa など）
    if (subCode && searchMap[subCode] !== undefined) {
      const idxSub = searchMap[subCode];
      if (kind === '取寄') results[idxSub][1] += qty;
      else results[idxSub][0] += qty;
    }

    // ② 親コードで検索している行（MAXM2511 など）
    if (searchMap[rowCode] !== undefined) {
      const idxParent = searchMap[rowCode];
      const hasVar = !!hasVariant[rowCode];
      const isBaseRow = !subCode || subCode === rowCode;

      // サフィックス付き行を持つ親コードの「合計行」は二重計上になるので無視
      if (hasVar && isBaseRow) {
        continue;
      }

      if (kind === '取寄') results[idxParent][1] += qty;
      else results[idxParent][0] += qty;
    }
  }

  // ===== B/C列に一括書き戻し =====
  sheet.getRange(2, 2, results.length, 2).setValues(results);

  Logger.log('在庫検索完了');

  // ★処理件数（検索対象コード数）を返す
  return searchKeys.length;
}

/****************************************
 * 条件付き書式だけ全部消したい時用（手動：UIあり）
 ****************************************/
function 条件付き書式を削除() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('在庫検索');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「在庫検索」シートが見つかりません');
    return;
  }

  sheet.clearConditionalFormatRules();

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).setBackground(null);
  }

  SpreadsheetApp.getUi().alert('条件付き書式と背景色を削除しました');
}
