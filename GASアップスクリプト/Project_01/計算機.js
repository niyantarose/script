/**
 * 仕入れ計算機（中国・韓国・台湾）
 *
 * ・サイドバー電卓: メニュー「💴 計算機」→「🧮 電卓を開く」
 *     国を選び、黄色の欄だけ入力 → おすすめ売値・うわのせ額・粗利率を自動表示。
 *     レートは自動取得（open.er-api.com）。盛りたい時は手入力で上書き可。
 * ・きれいな計算機シート: メニュー「💴 計算機」→「✨ きれいな計算機シートを作成」
 *     台湾計算機 / 韓国計算機 / 中国計算機 を色分けレイアウトで作成（元シートは残す）。
 *     シートのレートは GOOGLEFINANCE で自動。取得不可なら従来値にフォールバック。
 */

/* ===== サイドバー ===== */
function 計算機サイドバーを開く() {
  const html = HtmlService.createHtmlOutputFromFile('計算機サイドバー')
    .setTitle('仕入れ計算機');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** サイドバーから呼ぶ: 為替レートを自動取得（X→円） */
function 計算機_レート取得() {
  const fb = { ok: false, tw: 5.09, kr: 0.105, cn: 23.82, date: '' };
  try {
    const res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/JPY', { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return fb;
    const j = JSON.parse(res.getContentText());
    if (!j || !j.rates) return fb;
    const r3 = v => Math.round(v * 1000) / 1000;
    const r4 = v => Math.round(v * 10000) / 10000;
    return {
      ok: true,
      tw: r3(1 / j.rates.TWD),
      kr: r4(1 / j.rates.KRW),
      cn: r3(1 / j.rates.CNY),
      date: j.time_last_update_utc || '',
    };
  } catch (e) {
    return fb;
  }
}

/* ===== きれいな計算機シート ===== */
function 計算機_きれいなシートを作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // type: title / section / blank / input / auto / result
  const 台湾 = [
    ['台湾 仕入れ計算機', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色に入力）──', '', 'section'],
    ['本の価格(元)', 820, 'input', '0.0', '元'],
    ['目標粗利率', 0.25, 'input', '0%', '0.25=25%'],
    ['使用レート(円/元)', '=IFERROR(B7,5.09)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:TWDJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['送料(元)', '=B4*0.35', 'auto', '0.0', '本の価格×0.35'],
    ['原価(円)', '=ROUND((B4+B10)*B6,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['', '', 'blank'],
    ['── 結果（緑）──', '', 'section'],
    ['おすすめ売値', '=ROUND(B11/(1-B5-0.09-0.1),-2)', 'result', '¥#,##0', '←これで出品(100円単位)'],
    ['うわのせ額', '=B14-B11', 'result', '¥#,##0', '売値−原価'],
    ['実質粗利率', '=(B14-B14*(0.09+0.1)-B11)/B14', 'result', '0%', '20%以上が目安'],
    ['判定', '=IF(B16<0.2,"⚠ 20%割れ 注意","OK")', 'result', '@', ''],
  ];

  const 中国 = [
    ['中国 仕入れ計算機', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色に入力）──', '', 'section'],
    ['商品価格(元)', 37.4, 'input', '0.0', '元'],
    ['目標粗利率', 0.3, 'input', '0%', '0.3=30%'],
    ['使用レート(円/元)', '=IFERROR(B7,23.82)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:CNYJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['送料(元)', '=B4*0.3', 'auto', '0.0', '商品価格×0.3'],
    ['原価(円)', '=ROUND((B4+B10)*B6,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['', '', 'blank'],
    ['── 結果（緑）──', '', 'section'],
    ['おすすめ売値', '=ROUND(B11/(1-B5-0.12-0.1),-2)', 'result', '¥#,##0', '←これで出品(カード3%込)'],
    ['うわのせ額', '=B14-B11', 'result', '¥#,##0', '売値−原価'],
    ['実質粗利率', '=(B14-B14*(0.12+0.1)-B11)/B14', 'result', '0%', '20%以上が目安'],
    ['判定', '=IF(B16<0.2,"⚠ 20%割れ 注意","OK")', 'result', '@', ''],
  ];

  const 韓国 = [
    ['韓国 仕入れ計算機', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色に入力）──', '', 'section'],
    ['本の価格(ウォン)', 180900, 'input', '#,##0', 'ウォン'],
    ['重さ(g)', 300, 'input', '#,##0', '縦×横×厚さ(cm)×0.45 ≒ g'],
    ['目標粗利率', 0.25, 'input', '0%', '0.25=25%'],
    ['使用レート(円/ウォン)', '=IFERROR(B8,0.105)', 'input', '0.0000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:KRWJPY"),"取得不可")', 'auto', '0.0000', '自動取得'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['原価(円)', '=ROUND(B4*B7+B5*0.6,0)', 'auto', '¥#,##0', '価格×レート + 重さ×0.6'],
    ['', '', 'blank'],
    ['── 結果（緑）──', '', 'section'],
    ['おすすめ売値', '=ROUND(B11/(1-B6-0.09-0.1),-2)', 'result', '¥#,##0', '←これで出品(100円単位)'],
    ['うわのせ額', '=B14-B11', 'result', '¥#,##0', '売値−原価'],
    ['実質粗利率', '=(B14-B14*(0.09+0.1)-B11)/B14', 'result', '0%', '20%以上が目安'],
    ['判定', '=IF(B16<0.2,"⚠ 20%割れ 注意","OK")', 'result', '@', ''],
  ];

  計算機_buildSheet_(ss, '台湾計算機', 台湾);
  計算機_buildSheet_(ss, '韓国計算機', 韓国);
  計算機_buildSheet_(ss, '中国計算機', 中国);

  ss.toast('台湾計算機・韓国計算機・中国計算機 を作成しました（元シートはそのまま残しています）');
}

function 計算機_buildSheet_(ss, name, rows) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 300);

  const BG = { input: '#FFF2CC', auto: '#F1EFE8', result: '#D9EAD3' };
  let profitA1 = null;

  rows.forEach(function (r, idx) {
    const row = idx + 1;
    const label = r[0], val = r[1], type = r[2], fmt = r[3], note = r[4];
    if (type === 'blank') return;
    if (type === 'title') {
      sh.getRange(row, 1, 1, 3).merge();
      sh.getRange(row, 1).setValue(label).setFontSize(14).setFontWeight('bold');
      return;
    }
    if (type === 'section') {
      sh.getRange(row, 1, 1, 3).merge();
      sh.getRange(row, 1).setValue(label).setFontWeight('bold').setFontColor('#6B6B6B');
      return;
    }
    const a = sh.getRange(row, 1);
    const b = sh.getRange(row, 2);
    a.setValue(label);
    if (typeof val === 'string' && val.charAt(0) === '=') b.setFormula(val);
    else b.setValue(val);
    if (fmt) b.setNumberFormat(fmt);
    if (BG[type]) { a.setBackground(BG[type]); b.setBackground(BG[type]); }
    if (type === 'result' && label.indexOf('おすすめ') >= 0) b.setFontSize(14).setFontWeight('bold');
    if (note) sh.getRange(row, 3).setValue(note).setFontColor('#999999').setFontSize(10);
    if (label.indexOf('粗利率') >= 0 && label.indexOf('目標') < 0) profitA1 = b.getA1Notation();
  });

  if (profitA1) {
    const rng = sh.getRange(profitA1);
    const bad = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.2).setBackground('#F4C7C3').setFontColor('#CC0000').setBold(true)
      .setRanges([rng]).build();
    const ok = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.2).setBackground('#D9EAD3')
      .setRanges([rng]).build();
    sh.setConditionalFormatRules([bad, ok]);
  }
}
