/**
 * ②オプション入力シート：
 * 同じ親コード(A) + 同じ種類番号(サブコードCから抽出) の a/b ペアで
 * D列（オプション表示名）を揃える。
 *
 * 仕様（既定）:
 * - a を優先して b にコピー（bが空のとき）
 * - aが空で bにある場合は b→a で補完
 * - 両方に値がある場合は何もしない（設定で上書きも可能）
 *
 * 使い方:
 * - Apps Script に貼る
 * - 関数「②_バリエーション名をaとbで揃える」を実行
 */
function バリエーション名をaとbで揃える() {
  const CFG = {
    SHEET_2: '②オプション入力シート',
    HEADER_ROWS: 1,     // ②は1行目ヘッダー
    COL_CODE: 1,        // A 親コード
    COL_SUB: 3,         // C サブコード
    COL_NAME: 4,        // D 表示名

    // 上書き方針
    // "fill_only": 空だけ埋める（おすすめ）
    // "force_a_to_b": bを常にaで上書き（aがある時）
    // "force_b_to_a": aを常にbで上書き（bがある時）
    MODE: 'fill_only',

    // 空行（Aが空）の扱い：無視する
  };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_2);
  if (!sh) throw new Error('②オプション入力シート が見つかりません');

  const lastRow = sh.getLastRow();
  if (lastRow <= CFG.HEADER_ROWS) return;

  const numRows = lastRow - CFG.HEADER_ROWS;
  const range = sh.getRange(CFG.HEADER_ROWS + 1, 1, numRows, CFG.COL_NAME); // A:D
  const values = range.getValues(); // [A,B,C,D]

  // key = code::kindNo  -> { aRow, bRow, aName, bName }
  const map = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const code = String(row[CFG.COL_CODE - 1] || '').trim();
    const sub  = String(row[CFG.COL_SUB  - 1] || '').trim();
    const name = String(row[CFG.COL_NAME - 1] || '').trim();
    if (!code || !sub) continue;

    const ab = sub.slice(-1).toLowerCase(); // a or b を期待
    if (ab !== 'a' && ab !== 'b') continue;

    const kind = kindNoFromSub_(code, sub); // 1..N
    const key = `${code}::${kind}`;

    if (!map.has(key)) map.set(key, {});
    const obj = map.get(key);

    if (ab === 'a') {
      obj.aRow = i;   // values内のindex
      obj.aName = name;
    } else {
      obj.bRow = i;
      obj.bName = name;
    }
  }

  let changed = 0;

  // 変更は values を直接書き換えて最後に setValues で一括反映
  for (const [key, obj] of map.entries()) {
    if (obj.aRow == null || obj.bRow == null) continue; // a/b揃ってない

    const aName = String(obj.aName || '').trim();
    const bName = String(obj.bName || '').trim();

    if (CFG.MODE === 'fill_only') {
      if (aName && !bName) {
        values[obj.bRow][CFG.COL_NAME - 1] = aName;
        changed++;
      } else if (!aName && bName) {
        values[obj.aRow][CFG.COL_NAME - 1] = bName;
        changed++;
      }
      // 両方ある → 何もしない
      continue;
    }

    if (CFG.MODE === 'force_a_to_b') {
      if (aName && bName !== aName) {
        values[obj.bRow][CFG.COL_NAME - 1] = aName;
        changed++;
      }
      continue;
    }

    if (CFG.MODE === 'force_b_to_a') {
      if (bName && aName !== bName) {
        values[obj.aRow][CFG.COL_NAME - 1] = bName;
        changed++;
      }
      continue;
    }
  }

  if (changed > 0) {
    range.setValues(values);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `a/b の表示名を同期しました：${changed} 箇所`,
    '②同期',
    8
  );
}

/**
 * subCode から種類番号を推定
 * - code-3a / code-3b → 3
 * - codea / codeb     → 1
 */
function kindNoFromSub_(code, sub) {
  const esc = String(code).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp('^' + esc + '-(\\d+)[ab]$', 'i');
  const m = String(sub).match(re);
  if (m) return Number(m[1]);

  const re1 = new RegExp('^' + esc + '[ab]$', 'i');
  if (re1.test(String(sub))) return 1;

  return 1;
}
