/**
 * 仕入れ計算機 v2（中国・韓国・台湾）
 *
 * レイアウト: 入力 → 微調整結果 → おすすめ(理論/100円) → 自動計算 → 【別設定】原価帯(最下)
 * ・原価帯の利益率は一番下で設定 → 原価に応じて自動適用
 * ・黄色の目標粗利率は通常それに連動（上書き可）
 * ・実際のうわのせは黄色で手入力
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

/** 原価帯別目標粗利率（JS側・初期値用） */
function 計算機_原価帯粗利率_(cost) {
  const c = Number(cost) || 0;
  if (c <= 2000) return 0.35;
  if (c <= 6000) return 0.25;
  if (c <= 12000) return 0.2;
  if (c <= 20000) return 0.1;
  return 0.08;
}

/** 原価帯セルを参照する IFS 式 */
function 計算機_原価帯粗利率式_(costA1, r1, r2, r3, r4, r5) {
  return '=IFS(' +
    costA1 + '<=2000,' + r1 + ',' +
    costA1 + '<=6000,' + r2 + ',' +
    costA1 + '<=12000,' + r3 + ',' +
    costA1 + '<=20000,' + r4 + ',' +
    'TRUE,' + r5 + ')';
}

/** 100円丸め基準の推奨うわのせ（微調整欄の初期値用） */
function 計算機_推奨うわのせ100_(cost, goal, feeTotal) {
  const c = Number(cost) || 0;
  const g = Number(goal) || 0;
  const f = Number(feeTotal) || 0;
  const denom = 1 - g - f;
  if (denom <= 0 || c <= 0) return 0;
  const theory = c / denom;
  const round100 = Math.round(theory / 100) * 100;
  return Math.max(0, Math.round(round100 - c));
}

/** 【別設定】原価帯（シート最下部用） */
function 計算機_原価帯設定行_() {
  return [
    ['── 【別設定】原価帯別の目標粗利率 ──', '', 'section'],
    ['〜 2,000円まで', 0.35, 'band', '0%', 'ルール設定。原価がこの帯ならこの率を自動適用'],
    ['〜 6,000円まで', 0.25, 'band', '0%', 'ルール設定。原価がこの帯ならこの率を自動適用'],
    ['〜 12,000円まで', 0.2, 'band', '0%', 'ルール設定。原価がこの帯ならこの率を自動適用'],
    ['〜 20,000円まで', 0.1, 'band', '0%', 'ルール設定。原価がこの帯ならこの率を自動適用'],
    ['20,000円超', 0.08, 'band', '0%', 'ルール設定。原価がこの帯ならこの率を自動適用'],
  ];
}

function 計算機_きれいなシートを作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const twCost = Math.round((820 + 820 * 0.35) * 5.09);
  const twFee = 0.09 + 0.1;
  const twGoal = 計算機_原価帯粗利率_(twCost);
  const twUwa = 計算機_推奨うわのせ100_(twCost, twGoal, twFee);

  const cnCost = Math.round((37.4 + 37.4 * 0.3) * 23.82);
  const cnFee = 0.09 + 0.03 + 0.1;
  const cnGoal = 計算機_原価帯粗利率_(cnCost);
  const cnUwa = 計算機_推奨うわのせ100_(cnCost, cnGoal, cnFee);

  const krCost = Math.round(180900 * 0.105 + 300 * 0.6);
  const krFee = 0.09 + 0.1;
  const krGoal = 計算機_原価帯粗利率_(krCost);
  const krUwa = 計算機_推奨うわのせ100_(krCost, krGoal, krFee);

  // 台湾
  // 入力: B4-B11 / 微調整: B14-B17 / おすすめ: B20-B21
  // 自動: 手数料B24 送料B25 原価B26 適用B27 / 帯: B30-B34
  const 台湾 = [
    ['台湾 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['本の価格(元)', 820, 'input', '0.0', '元'],
    ['使用レート(円/元)', '=IFERROR(B6,5.09)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:TWDJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.35, 'input', '0%', '送料(元)=本の価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', '=IFERROR(B27,' + twGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', twUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(100円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B26+B11', 'actual', '¥#,##0', '式: 原価(B26)+うわのせ(B11)'],
    ['実質粗利額', '=B14-B14*B24-B26', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B15/B14,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B16<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['理論売値', '=IFERROR(ROUND(B26/(1-B10-B24),0),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計) 例:8480'],
    ['おすすめ売値(100円丸め)', '=ROUND(B20,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め 例:8500'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9', 'auto', '0%', 'モール+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '本の価格×送料率'],
    ['原価(円)', '=ROUND((B4+B25)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B26', 'B30', 'B31', 'B32', 'B33', 'B34'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 中国
  // 入力: B4-B12 / 微調整: B15-B18 / おすすめ: B21-B22
  // 自動: 手数料B25 送料B26 原価B27 適用B28 / 帯: B31-B35
  const 中国 = [
    ['中国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['商品価格(元)', 37.4, 'input', '0.0', '元'],
    ['使用レート(円/元)', '=IFERROR(B6,23.82)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:CNYJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.3, 'input', '0%', '送料(元)=商品価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['カード手数料', 0.03, 'input', '0%', '決済手数料'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', '=IFERROR(B28,' + cnGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', cnUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(100円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B27+B12', 'actual', '¥#,##0', '式: 原価(B27)+うわのせ(B12)'],
    ['実質粗利額', '=B15-B15*B25-B27', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B16/B15,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B17<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['理論売値', '=IFERROR(ROUND(B27/(1-B11-B25),0),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計) 例:8480'],
    ['おすすめ売値(100円丸め)', '=ROUND(B21,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め 例:8500'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9+B10', 'auto', '0%', 'モール+カード+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '商品価格×送料率'],
    ['原価(円)', '=ROUND((B4+B26)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B27', 'B31', 'B32', 'B33', 'B34', 'B35'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 韓国
  // 入力: B4-B12 / 微調整: B15-B18 / おすすめ: B21-B22
  // 自動: 手数料B25 原価B26 適用B27 / 帯: B30-B34
  const 韓国 = [
    ['韓国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['本の価格(ウォン)', 180900, 'input', '#,##0', 'ウォン'],
    ['重さ(g)', 300, 'input', '#,##0', 'グラムで重さを入力'],
    ['使用レート(円/ウォン)', '=IFERROR(B7,0.105)', 'input', '0.0000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:KRWJPY"),"取得不可")', 'auto', '0.0000', '自動取得'],
    ['重量係数(送料)', 0.6, 'input', '0.0', '送料(円)=重さ×係数'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', '=IFERROR(B27,' + krGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', krUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(100円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B26+B12', 'actual', '¥#,##0', '式: 原価(B26)+うわのせ(B12)'],
    ['実質粗利額', '=B15-B15*B25-B26', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B16/B15,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B17<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['理論売値', '=IFERROR(ROUND(B26/(1-B11-B25),0),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計) 例:8480'],
    ['おすすめ売値(100円丸め)', '=ROUND(B21,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め 例:8500'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B9+B10', 'auto', '0%', 'モール+消費税'],
    ['原価(円)', '=ROUND(B4*B6+B5*B8,0)', 'auto', '¥#,##0', '価格×レート+重さ×係数'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B26', 'B30', 'B31', 'B32', 'B33', 'B34'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  計算機_buildSheet_(ss, '台湾計算機', 台湾);
  計算機_buildSheet_(ss, '韓国計算機', 韓国);
  計算機_buildSheet_(ss, '中国計算機', 中国);

  ss.toast('台湾・韓国・中国 計算機 v2 を再作成しました');
}

function 計算機_setNote_(range, note) {
  if (!note) return;
  const text = String(note);
  if (text.charAt(0) === '=') range.setValue("'" + text);
  else range.setValue(text);
  range.setFontColor('#5F6368').setFontSize(11);
}

function 計算機_buildSheet_(ss, name, rows) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 400);
  sh.getRange(1, 1, 80, 3).setFontSize(12);

  const BG = {
    input: '#FFF2CC',
    band: '#FCE8B2',
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
      sh.getRange(row, 1).setValue(label).setFontSize(18).setFontWeight('bold');
      return;
    }
    if (type === 'section') {
      sh.getRange(row, 1, 1, 3).merge();
      sh.getRange(row, 1).setValue(label).setFontWeight('bold').setFontSize(13).setFontColor('#3C4043');
      return;
    }
    const a = sh.getRange(row, 1);
    const b = sh.getRange(row, 2);
    a.setValue(label).setFontSize(13);
    if (typeof val === 'string' && val.charAt(0) === '=') b.setFormula(val);
    else if (val === '' && type === 'input') b.setValue('');
    else b.setValue(val);
    if (fmt) b.setNumberFormat(fmt);
    b.setFontSize(14);
    if (BG[type]) {
      a.setBackground(BG[type]);
      b.setBackground(BG[type]);
    }
    if (type === 'result' && (label.indexOf('おすすめ') >= 0 || label.indexOf('理論') >= 0)) {
      b.setFontSize(16).setFontWeight('bold');
    }
    if (type === 'actual' && label === '実際売値') {
      b.setFontSize(16).setFontWeight('bold');
    }
    if (type === 'band') {
      a.setFontSize(13);
      b.setFontSize(14).setFontWeight('bold');
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
