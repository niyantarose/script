/**
 * 指定した code の Yahoo options 文字列（D列用）を作成する
 *  - ②オプション入力シート A:D を使用
 *  - オプション入力シートに行が無いときは "" を返す
 *    （その場合は呼び出し側で「a/b のシンプル形」を自動生成）
 *
 * @param {string} code 親コード（①商品入力シート C列）
 * @return {string} Yahoo CSV 用 options 文字列（D列）
 */
function 作成Yahooオプション文字列_(code, otoriLabel) {
  otoriLabel = otoriLabel || 'お取り寄せ（韓国から）';

  const ss = SpreadsheetApp.getActive();
  const optSheet = ss.getSheetByName('②オプション入力シート');
  if (!optSheet) throw new Error('②オプション入力シート が見つかりません。');

  const lastRow = optSheet.getLastRow();
  if (lastRow < 2) return '';

  const values = optSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = values.filter(r =>
    String(r[0]).trim() === String(code).trim() &&
    String(r[2]).trim() !== ''
  );
  if (rows.length === 0) return '';

  const groups = {};
  rows.forEach(r => {
    const subCode = String(r[2]).trim();
    const name    = String(r[3]).trim();
    if (!subCode) return;

    const m = subCode.match(/-(\d+)[ab]$/);
    const baseKey = m ? m[1] : subCode.slice(0, -1);

    if (!groups[baseKey]) groups[baseKey] = { name, items: [] };
    groups[baseKey].items.push(subCode);
  });

  const baseKeys = Object.keys(groups).sort((a,b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a).localeCompare(String(b));
  });

  const 即納パーツ = [];
  const 取寄パーツ = [];
  let 即納番号 = 1;
  let 取寄番号 = 1;

  baseKeys.forEach(key => {
    const group = groups[key];
    const label = group.name || (code + '-' + key);

    let subA = '';
    let subB = '';
    group.items.forEach(subCode => {
      if (subCode.endsWith('a')) subA = subCode;
      else if (subCode.endsWith('b')) subB = subCode;
    });

    if (subA) {
      即納パーツ.push(
        '★在庫の設定:即納（日本在庫）' +
        '#★種類の選択:' + 即納番号 + '.' + label +
        '=' + subA
      );
      即納番号++;
    }

    if (subB) {
      取寄パーツ.push(
        '★在庫の設定:' + otoriLabel +
        '#★種類の選択:' + 取寄番号 + '.' + label +
        '=' + subB
      );
      取寄番号++;
    }
  });

  return 即納パーツ.concat(取寄パーツ).join('&');
}


/**
 * F列（options / オプション項目名）用の文字列を作る
 *  - ②オプション入力シート A:D を使う
 *  - サブコードが 1 組（a/b）だけのときも複数のときもここでまとめて作る
 *
 * @param {string} code 親コード（①商品入力シート C列）
 * @return {string} F列用の文字列
 *   例：
 *   ★在庫の選択 即納（日本在庫） お取り寄せ（韓国から）
 *   ★種類の選択 1.TEST03-01 2.TEST03-02 3.TEST03-03
 */
/**
 * F列(options / オプション項目名)用の文字列を作る
 *  - ②オプション入力シート A:D を使う
 *  - サブコードが 1 組(a/b)だけのときも複数のときもここでまとめて作る
 *
 * @param {string} code 親コード(①商品入力シート C列)
 * @return {string} F列用の文字列
 *   例:
 *   ★在庫の選択 即納(日本在庫) お取り寄せ(韓国から)
 *   [空白行]
 *   ★種類の選択 1.TEST03-01 2.TEST03-02 3.TEST03-03
 */
function 作成Yahooオプション項目名_(code, otoriLabel) {
  otoriLabel = otoriLabel || 'お取り寄せ（韓国から）';

  const ss = SpreadsheetApp.getActive();
  const optSheet = ss.getSheetByName('②オプション入力シート');
  if (!optSheet) throw new Error('②オプション入力シート が見つかりません。');

  const lastRow = optSheet.getLastRow();
  const 第一行 = '★在庫の選択 即納（日本在庫） ' + otoriLabel;

  if (lastRow < 2) return 第一行;

  const values = optSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = values.filter(r =>
    String(r[0]).trim() === String(code).trim() &&
    String(r[2]).trim() !== ''
  );
  if (rows.length === 0) return 第一行;

  const groups = {};
  rows.forEach(r => {
    const subCode = String(r[2]).trim();
    const name    = String(r[3]).trim();
    if (!subCode) return;

    const m = subCode.match(/-(\d+)[ab]$/);
    const baseKey = m ? m[1] : subCode.slice(0, -1);

    if (!groups[baseKey]) groups[baseKey] = name || (code + '-' + baseKey);
  });

  const baseKeys = Object.keys(groups).sort((a,b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a).localeCompare(String(b));
  });

  const labels = baseKeys.map((key, idx) => (idx + 1) + '.' + (groups[key] || (code + '-' + key)));
  const 第二行 = '★種類の選択 ' + labels.join(' ');

  return 第一行 + '\n\n' + 第二行;
}
