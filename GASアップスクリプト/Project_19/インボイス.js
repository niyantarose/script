// ============================================================
//  EMS大邱作業データ → 積荷NOW（同じGoogleスプレッドシート内）転記
// ============================================================

const EMS_SHEET  = 'EMS大邱作業データ';
const DEST_SHEET = '積荷NOW';

// EMSリスト側の列番号（実測値）
const COL_DESC        = 6;   // F列: Description / Title
const COL_ITEM        = 8;   // H列: ItemCode
const COL_QTY         = 9;   // I列: Qty
const COL_TYPE        = 11;  // K列: Type
const COL_WEIGHT      = 12;  // L列: weight(g)
const COL_BILLED_UNIT = 18;  // R列: Billed Unit Price(円)
const COL_BILLED_AMT  = 19;  // S列: Billed Amount(円)

// 積荷NOW側ヘッダー（B列・E列は空列）
const HEADERS = [
  'Description / Title', '列1', 'ItemCode', 'Qty', '列2', 'Type', 'weight(g)', 'Unit Price'
];

// 積荷NOWの書き込み列（1始まり）
const DST_A = 1;  // A列: Description / Title
// B列(2): 列1 → 空
const DST_C = 3;  // C列: ItemCode
const DST_D = 4;  // D列: Qty
// E列(5): 列2 → 空
const DST_F = 6;  // F列: Type
const DST_G = 7;  // G列: weight(g)
const DST_H = 8;  // H列: Unit Price(円)
// Amount(円)はExcelマクロが計算するので不要

const DST_HEADER_ROW = 3;  // ヘッダー行
const DST_DATA_START = 4;  // データ開始行
// ================


function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('📦 積荷NOW 転記')
    .setWidth(260);
  SpreadsheetApp.getUi().showSidebar(html);
}

function インボイスメニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('インボイス作成')
    .addItem('📦 積荷NOW 転記サイドバー', 'showSidebar')
    .addToUi();
}

function transferToIntegration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const emsSheet = ss.getSheetByName(EMS_SHEET);
  if (!emsSheet) {
    return { ok: false, msg: `シート「${EMS_SHEET}」が見つかりません。` };
  }

  const selection = emsSheet.getActiveRange();
  if (!selection) {
    return { ok: false, msg: '行を選択してから転記ボタンを押してください。' };
  }

  const startRow = selection.getRow();
  const numRows  = selection.getNumRows();

  if (startRow <= 2) {
    return { ok: false, msg: 'ヘッダー行は選択しないでください。3行目以降を選択してください。' };
  }

  // 転記データを収集
  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const r = startRow + i;
    const desc   = emsSheet.getRange(r, COL_DESC).getValue();
    const item   = emsSheet.getRange(r, COL_ITEM).getValue();
    const qty    = emsSheet.getRange(r, COL_QTY).getValue();
    const type   = emsSheet.getRange(r, COL_TYPE).getValue();
    const weight = emsSheet.getRange(r, COL_WEIGHT).getValue();
    const bUnit  = emsSheet.getRange(r, COL_BILLED_UNIT).getValue();

    if (!desc && !item) continue;

    rows.push({ desc, item, qty, type, weight, bUnit });
  }

  if (rows.length === 0) {
    return { ok: false, msg: '転記できるデータがありませんでした。' };
  }

  // 転記先シート（なければ作成）
  let destSheet = ss.getSheetByName(DEST_SHEET);
  if (!destSheet) {
    destSheet = ss.insertSheet(DEST_SHEET);
  }

  // クリア＆ヘッダー書き込み
  destSheet.clearContents();
  destSheet.getRange(DST_HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);

  // データ書き込み（列ごとに正確に）
  rows.forEach((row, idx) => {
    const r = DST_DATA_START + idx;
    destSheet.getRange(r, DST_A).setValue(row.desc);
    destSheet.getRange(r, DST_C).setValue(row.item);
    destSheet.getRange(r, DST_D).setValue(row.qty);
    destSheet.getRange(r, DST_F).setValue(row.type);
    destSheet.getRange(r, DST_G).setValue(row.weight);
    destSheet.getRange(r, DST_H).setValue(row.bUnit);
  });

  SpreadsheetApp.flush();

  return { ok: true, msg: `✅ ${rows.length}件を「積荷NOW」に転記しました！` };
}


function getSidebarHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; margin: 0; padding: 16px; background: #f8f9fa; color: #333; }
    h2 { font-size: 15px; margin: 0 0 6px 0; color: #1a1a2e; }
    .sub { font-size: 11px; color: #888; margin-bottom: 20px; line-height: 1.5; }
    .step { background: #fff; border-radius: 8px; padding: 12px 14px; margin-bottom: 12px; border-left: 3px solid #4a90d9; font-size: 12px; line-height: 1.6; }
    .step-num { font-weight: bold; color: #4a90d9; }
    .btn { width: 100%; padding: 13px; background: #2d6cdf; color: white; border: none; border-radius: 8px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 4px; }
    .btn:hover { background: #1a55c4; }
    .btn:disabled { background: #aaa; cursor: not-allowed; }
    .msg { margin-top: 14px; padding: 10px 12px; border-radius: 6px; font-size: 12px; line-height: 1.5; display: none; }
    .msg.ok  { background: #e6f4ea; color: #1e6b3a; border: 1px solid #a8d5b5; }
    .msg.err { background: #fdecea; color: #a61c00; border: 1px solid #f5c6c6; }
    .loading { display: none; font-size: 12px; color: #888; margin-top: 10px; text-align: center; }
  </style>
</head>
<body>
  <h2>📦 積荷NOW 転記</h2>
  <p class="sub">EMSリストの行を選択して<br>下のボタンで転記します</p>
  <div class="step"><span class="step-num">① </span>「EMS大邱作業データ」シートで<br>転記したい行を選択（複数行OK）</div>
  <div class="step"><span class="step-num">② </span>ボタンを押すと<br>「積荷NOW」シートに転記されます<br><span style="color:#e05;">※既存データは上書きされます</span></div>
  <button class="btn" id="btnTransfer" onclick="doTransfer()">転記する</button>
  <div class="loading" id="loading">⏳ 転記中...</div>
  <div class="msg" id="msg"></div>
  <script>
    function doTransfer() {
      const btn = document.getElementById('btnTransfer');
      const loading = document.getElementById('loading');
      const msg = document.getElementById('msg');
      btn.disabled = true;
      loading.style.display = 'block';
      msg.style.display = 'none';
      google.script.run
        .withSuccessHandler(function(res) {
          btn.disabled = false;
          loading.style.display = 'none';
          msg.className = 'msg ' + (res.ok ? 'ok' : 'err');
          msg.textContent = res.msg;
          msg.style.display = 'block';
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          loading.style.display = 'none';
          msg.className = 'msg err';
          msg.textContent = '❌ エラー: ' + err.message;
          msg.style.display = 'block';
        })
        .transferToIntegration();
    }
  </script>
</body>
</html>
  `;
}
