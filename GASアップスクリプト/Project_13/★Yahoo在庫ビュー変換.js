function Yahoo在庫ビューを更新_フィルター維持_SheetsAPI() {
  console.log('=== 処理開始 (Sheets API版・最適化) ===');
  const startTime = new Date();

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  const SRC_NAME = 'Yahoo在庫取込';
  const DST_NAME = 'Yahoo在庫ビュー';

  const DATA_START_ROW   = 3;
  const OUTPUT_COL_WIDTH = 7;   // A〜G
  const TAKE_COLS        = 4;   // 取込は A〜D
  const CHUNK_SIZE       = 30000;

  // ==========================================
  // Sheets API でデータ読み込み（超高速）
  // ==========================================
  const range = `'${SRC_NAME}'!A:${colLetter_(TAKE_COLS)}`;
  console.log(`読み込み範囲: ${range}`);

  let srcData;
  try {
    const result = Sheets.Spreadsheets.Values.get(ssId, range, {
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });
    srcData = result.values || [];
  } catch (e) {
    console.error('データ読み込みエラー:', e);
    return;
  }

  console.log(`データ読み込み完了: ${srcData.length}行`);

  if (srcData.length < 2) {
    ss.toast('データなし', '完了', 3);
    return;
  }

  // ヘッダー判定
  const a1 = String(srcData[0][0] || '').toLowerCase();
  const startIdx = (!a1 || a1 === 'code') ? 1 : 0;

  // ==========================================
  // グルーピング（メモリ内）
  // ==========================================
  const nameByCode = new Map();
  const rowsByCode = new Map();

  for (let r = startIdx; r < srcData.length; r++) {
    const row  = srcData[r] || [];
    const code = String(row[0] ?? '').trim();
    if (!code) continue;

    const name = String(row[1] ?? '').trim();
    const sub  = String(row[2] ?? '').trim();
    const qtyRaw = row[3];
    const qty  = (qtyRaw === '' || qtyRaw == null) ? 0 : (Number(qtyRaw) || 0);

    if (!nameByCode.has(code) && name) nameByCode.set(code, name);

    let suffixType = 'sokuno';
    if (sub) {
      const lastChar = sub.slice(-1).toLowerCase();
      if (lastChar === 'b') suffixType = 'otoriyose';
    }

    if (!rowsByCode.has(code)) rowsByCode.set(code, []);
    rowsByCode.get(code).push({ code, name, sub, qty, suffixType });
  }

  console.log(`グルーピング完了: ${rowsByCode.size}件`);

  // ==========================================
  // 出力データ作成
  // ==========================================
  const out = [];
  const warnRows = []; // 0-based（outの行index）
  const codes = Array.from(rowsByCode.keys()).sort();

  for (let i = 0; i < codes.length; i++) {
    const code  = codes[i];
    const group = rowsByCode.get(code) || [];
    const baseName = nameByCode.get(code) || (group[0] ? group[0].name : '');

    const hasSub = group.some(x => x.sub);

    if (hasSub) {
      const children = group
        .filter(x => x.sub)
        .sort((a, b) => String(a.sub).localeCompare(String(b.sub)));

      for (const it of children) {
        const built = buildUnionKeyWithWarn_(it.code, it.sub);

        out.push([
          it.code,
          it.name || baseName,
          it.sub,
          it.suffixType === 'sokuno' ? it.qty : 0,
          it.suffixType === 'otoriyose' ? it.qty : 0,
          built.key,
          '子'
        ]);

        if (built.warn) warnRows.push(out.length - 1);
      }

    } else {
      let totalSokuno = 0, totalOtoriyose = 0;
      for (const it of group) {
        if (it.suffixType === 'sokuno') totalSokuno += it.qty;
        else totalOtoriyose += it.qty;
      }

      out.push([
        code,
        baseName,
        '',
        totalSokuno,
        totalOtoriyose,
        code,
        '親'
      ]);
    }
  }

  console.log(`出力データ作成完了: ${out.length}行, 警告行: ${warnRows.length}件`);

  // ==========================================
  // Sheets API で書き込み（超高速）
  // ==========================================
  const totalRows = out.length;
  const dstSheet = ss.getSheetByName(DST_NAME);
  const dstSheetId = dstSheet.getSheetId();

  if (totalRows > 0) {
    for (let i = 0; i < totalRows; i += CHUNK_SIZE) {
      const chunk = out.slice(i, Math.min(i + CHUNK_SIZE, totalRows));
      const startRow = DATA_START_ROW + i;
      const endRow   = startRow + chunk.length - 1;
      const writeRange = `'${DST_NAME}'!A${startRow}:${colLetter_(OUTPUT_COL_WIDTH)}${endRow}`;

      Sheets.Spreadsheets.Values.update(
        { values: chunk, majorDimension: 'ROWS' },
        ssId,
        writeRange,
        { valueInputOption: 'RAW' }
      );

      console.log(`書き込み: ${i + 1}〜${i + chunk.length}行目 完了`);
    }
  }

  // ==========================================
  // ★ Sheets API batchUpdate で色付け（超高速化）
  // ==========================================
  const colorStartTime = new Date();
  
  const requests = [];
  
  // F列全体の背景色をリセット（1リクエスト）
  requests.push({
    repeatCell: {
      range: {
        sheetId: dstSheetId,
        startRowIndex: DATA_START_ROW - 1,  // 0-based
        endRowIndex: DATA_START_ROW - 1 + totalRows,
        startColumnIndex: 5,  // F列 = 0-based で 5
        endColumnIndex: 6
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 1, green: 1, blue: 1 }  // 白
        }
      },
      fields: 'userEnteredFormat.backgroundColor'
    }
  });

  // 警告行をまとめて黄色に（連続行をマージして効率化）
  if (warnRows.length > 0) {
    // 連続する行をグループ化してリクエスト数を削減
    const ranges = mergeConsecutiveRows_(warnRows, DATA_START_ROW);
    
    for (const r of ranges) {
      requests.push({
        repeatCell: {
          range: {
            sheetId: dstSheetId,
            startRowIndex: r.start - 1,  // 0-based
            endRowIndex: r.end,          // exclusive
            startColumnIndex: 5,
            endColumnIndex: 6
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 1, green: 0.902, blue: 0.6 }  // #FFE699
            }
          },
          fields: 'userEnteredFormat.backgroundColor'
        }
      });
    }
  }

  // batchUpdate 実行（1回のAPI呼び出しで全色付け完了）
  if (requests.length > 0) {
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
      const colorElapsed = (new Date() - colorStartTime) / 1000;
      console.log(`色付け完了: ${colorElapsed.toFixed(2)}秒 (${requests.length}リクエスト)`);
    } catch (e) {
      console.error('色付けエラー:', e);
    }
  }

  // ==========================================
  // 余分な行をクリア（行が減った場合のみ）
  // ==========================================
  const newLastRow = DATA_START_ROW + out.length - 1;
  const currentLastRow = dstSheet.getLastRow();

  if (currentLastRow > newLastRow) {
    const clearRange = `'${DST_NAME}'!A${newLastRow + 1}:${colLetter_(OUTPUT_COL_WIDTH)}${currentLastRow}`;
    console.log(`余分な行をクリア: ${newLastRow + 1}〜${currentLastRow}行目`);
    Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);
    
    // 余分な行の背景色もクリア
    requests.length = 0;
    requests.push({
      repeatCell: {
        range: {
          sheetId: dstSheetId,
          startRowIndex: newLastRow,
          endRowIndex: currentLastRow,
          startColumnIndex: 5,
          endColumnIndex: 6
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 1, green: 1, blue: 1 }
          }
        },
        fields: 'userEnteredFormat.backgroundColor'
      }
    });
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
  }

  const elapsed = (new Date() - startTime) / 1000;
  console.log(`=== メイン処理完了: ${elapsed}秒 ===`);
  ss.toast(`Yahoo在庫ビュー更新完了：${out.length}行 (${elapsed.toFixed(1)}秒)`, '完了', 5);

  // ==========================================
  // 連続処理
  // ==========================================
  if (out.length > 0) {
    try {
      console.log('Amazon在庫照合 開始...');
      Amazon在庫照合();
      console.log('Amazon在庫照合 完了');
    } catch (e) {
      console.error('Amazon在庫照合エラー:', e);
    }

    try {
      console.log('Qoo10在庫照合 開始...');
      Qoo10在庫照合();
      console.log('Qoo10在庫照合 完了');
    } catch (e) {
      console.error('Qoo10在庫照合エラー:', e);
    }
  }

  const totalElapsed = (new Date() - startTime) / 1000;
  console.log(`=== 全処理完了: ${totalElapsed}秒 ===`);
}


/**
 * 連続する行インデックスをマージしてリクエスト数を削減
 * @param {number[]} rows - 0-based の出力配列インデックス
 * @param {number} dataStartRow - データ開始行（1-based）
 * @returns {Array<{start: number, end: number}>} - 1-based の行範囲
 */
function mergeConsecutiveRows_(rows, dataStartRow) {
  if (rows.length === 0) return [];
  
  const sorted = [...rows].sort((a, b) => a - b);
  const ranges = [];
  
  let rangeStart = sorted[0];
  let rangeEnd = sorted[0];
  
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === rangeEnd + 1) {
      // 連続している
      rangeEnd = sorted[i];
    } else {
      // 連続が途切れた → 保存して新しい範囲開始
      ranges.push({
        start: dataStartRow + rangeStart,
        end: dataStartRow + rangeEnd + 1  // exclusive
      });
      rangeStart = sorted[i];
      rangeEnd = sorted[i];
    }
  }
  
  // 最後の範囲を追加
  ranges.push({
    start: dataStartRow + rangeStart,
    end: dataStartRow + rangeEnd + 1
  });
  
  return ranges;
}


// 列番号→A1文字
function colLetter_(colNum) {
  let letter = '';
  while (colNum > 0) {
    const mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return letter;
}


// ★統合キー生成（警告判定付き）
function buildUnionKeyWithWarn_(code, sub) {
  const c = String(code || '').trim();
  const s = String(sub  || '').trim();
  if (!c) return { key: '', warn: true, reason: 'code空' };
  if (!s) return { key: c, warn: false, reason: '' };

  const cN = normalizeKeyPart_(c);
  const sN = normalizeKeyPart_(s);

  // subがcodeで始まるなら、そのままsub（＝二重防止）
  if (sN.startsWith(cN)) {
    return { key: s, warn: false, reason: '' };
  }

  // subがcodeで始まらない場合、末尾バリエーションだけ拾う（警告）
  const m = s.match(/(-\d+)?[A-Z]?[ab]$/);
  if (m) {
    return { key: c + m[0], warn: true, reason: 'subがcodeで始まらない' };
  }

  // どうにもならない → concat
  return { key: c + s, warn: true, reason: 'sub解析不能' };
}


function normalizeKeyPart_(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, '')
    .replace(/\u3000/g, '')
    .toUpperCase();
}