/**
 * 仕入れ計算機 v2（中国・韓国・台湾）
 *
 * ・おすすめ（目標から）と微調整（実際）を別セクションで表示
 * ・10円丸め / 100円丸めの両方
 * ・手数料・送料率・消費税は黄色欄で変更可
 */

/* ===== サイドバー ===== */
function 計算機サイドバーを開く() {
  const html = HtmlService.createHtmlOutputFromFile('計算機サイドバー')
    .setTitle('仕入れ計算機');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getKeisanRates() {
  return 計算機_レート取得();
}

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

/** 10円丸め基準の推奨うわのせ（微調整欄の初期値用） */
function 計算機_推奨うわのせ10_(cost, goal, feeTotal) {
  const c = Number(cost) || 0;
  const g = Number(goal) || 0;
  const f = Number(feeTotal) || 0;
  const denom = 1 - g - f;
  if (denom <= 0 || c <= 0) return 0;
  const theory = c / denom;
  const round10 = Math.round(theory / 10) * 10;
  return Math.max(0, Math.round(round10 - c));
}

function 計算機_きれいなシートを作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const twCost = Math.round((820 + 820 * 0.35) * 5.09);
  const twFee = 0.09 + 0.1;
  const twUwa = 計算機_推奨うわのせ10_(twCost, 0.25, twFee);

  const cnCost = Math.round((37.4 + 37.4 * 0.3) * 23.82);
  const cnFee = 0.09 + 0.03 + 0.1;
  const cnUwa = 計算機_推奨うわのせ10_(cnCost, 0.3, cnFee);

  const krCost = Math.round(180900 * 0.105 + 300 * 0.6);
  const krFee = 0.09 + 0.1;
  const krUwa = 計算機_推奨うわのせ10_(krCost, 0.25, krFee);

  const 台湾 = [
    ['台湾 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['本の価格(元)', 820, 'input', '0.0', '元'],
    ['目標粗利率', 0.25, 'input', '0%', '0.25=25%'],
    ['使用レート(円/元)', '=IFERROR(B7,5.09)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:TWDJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.35, 'input', '0%', '送料(元)=本の価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B9+B10', 'auto', '0%', 'モール+消費税 → 粗利計算に使用'],
    ['送料(元)', '=B4*B8', 'auto', '0.0', '本の価格×送料率'],
    ['原価(円)', '=ROUND((B4+B14)*B6,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B15/(1-B5-B13),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B18,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B18,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B19-B15', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B20-B15', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', twUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B15+B25', 'actual', '¥#,##0', '式: 原価(B15)+うわのせ(B25)'],
    ['実質粗利額', '=B26-B26*B13-B15', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B27)/B26', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B28<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ];

  const 中国 = [
    ['中国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['商品価格(元)', 37.4, 'input', '0.0', '元'],
    ['目標粗利率', 0.3, 'input', '0%', '0.3=30%'],
    ['使用レート(円/元)', '=IFERROR(B7,23.82)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:CNYJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.3, 'input', '0%', '送料(元)=商品価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['カード手数料', 0.03, 'input', '0%', '決済手数料'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B9+B10+B11', 'auto', '0%', 'モール+カード+消費税'],
    ['送料(元)', '=B4*B8', 'auto', '0.0', '商品価格×送料率'],
    ['原価(円)', '=ROUND((B4+B15)*B6,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B16/(1-B5-B14),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B19,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B19,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B20-B16', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B21-B16', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', cnUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B16+B26', 'actual', '¥#,##0', '式: 原価(B16)+うわのせ(B26)'],
    ['実質粗利額', '=B27-B27*B14-B16', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B28)/B27', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B29<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ];

  const 韓国 = [
    ['韓国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['本の価格(ウォン)', 180900, 'input', '#,##0', 'ウォン'],
    ['重さ(g)', 300, 'input', '#,##0', '縦×横×厚さ(cm)×0.45 ≒ g'],
    ['目標粗利率', 0.25, 'input', '0%', '0.25=25%'],
    ['使用レート(円/ウォン)', '=IFERROR(B8,0.105)', 'input', '0.0000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:KRWJPY"),"取得不可")', 'auto', '0.0000', '自動取得'],
    ['重量係数(送料)', 0.6, 'input', '0.0', '送料(円)=重さ×係数'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B10+B11', 'auto', '0%', 'モール+消費税'],
    ['原価(円)', '=ROUND(B4*B7+B5*B9,0)', 'auto', '¥#,##0', '価格×レート+重さ×係数'],
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B15/(1-B6-B14),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B18,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B18,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B19-B15', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B20-B15', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', krUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B15+B25', 'actual', '¥#,##0', '式: 原価(B15)+うわのせ(B25)'],
    ['実質粗利額', '=B26-B26*B14-B15', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B27)/B26', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B28<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ];

  計算機_buildSheet_(ss, '台湾計算機', 台湾);
  計算機_buildSheet_(ss, '韓国計算機', 韓国);
  計算機_buildSheet_(ss, '中国計算機', 中国);

  ss.toast('台湾・韓国・中国 計算機 v2 を再作成しました（おすすめ/実際を分離）');
}

function 計算機_setNote_(range, note) {
  if (!note) return;
  const text = String(note);
  // 先頭が = だとスプレッドシートが数式と解釈して #NAME? になるため文字列として書く
  if (text.charAt(0) === '=') range.setValue("'" + text);
  else range.setValue(text);
  range.setFontColor('#666666').setFontSize(10);
}

function 計算機_buildSheet_(ss, name, rows) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 360);

  const BG = {
    input: '#FFF2CC',
    auto: '#F1EFE8',
    result: '#D9EAD3',
    actual: '#D0E0F0',
  };
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
    if (type === 'result' && (label.indexOf('おすすめ') >= 0 || label.indexOf('理論') >= 0)) {
      b.setFontSize(13).setFontWeight('bold');
    }
    if (type === 'actual' && label === '実際売値') {
      b.setFontSize(13).setFontWeight('bold');
    }
    if (note) 計算機_setNote_(sh.getRange(row, 3), note);
    if (label === '実質粗利率' && type === 'actual') profitA1 = b.getA1Notation();
  });

  if (profitA1) {
    const rng = sh.getRange(profitA1);
    const bad = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.2).setBackground('#F4C7C3').setFontColor('#CC0000').setBold(true)
      .setRanges([rng]).build();
    const ok = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.2).setBackground('#D0E0F0')
      .setRanges([rng]).build();
    sh.setConditionalFormatRules([bad, ok]);
  }
}
