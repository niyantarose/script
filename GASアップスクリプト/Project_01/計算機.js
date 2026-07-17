/**
 * 仕入れ計算機 v2（中国・韓国・台湾）
 *
 * レイアウト: 入力 → 微調整 → おすすめ(10円/100円) → 自動価格計算結果(原価帯) → 【別設定】原価帯(最下)
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
  const twUwa = 計算機_推奨うわのせ10_(twCost, twGoal, twFee);

  const cnCost = Math.round((37.4 + 37.4 * 0.3) * 23.82);
  const cnFee = 0.09 + 0.03 + 0.1;
  const cnGoal = 計算機_原価帯粗利率_(cnCost);
  const cnUwa = 計算機_推奨うわのせ10_(cnCost, cnGoal, cnFee);

  const krCost = Math.round(180900 * 0.105 + 300 * 0.6);
  const krFee = 0.09 + 0.1;
  const krGoal = 計算機_原価帯粗利率_(krCost);
  const krUwa = 計算機_推奨うわのせ10_(krCost, krGoal, krFee);

  // 台湾
  // 入力 B4-B11 / 微調整 B14-B17 / おすすめ B20-B21 / 原価帯売値 B24-B25
  // 自動計算: 手数料B28 送料B29 原価B30 適用B31 / 帯 B34-B38
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
    ['目標粗利率', '=IFERROR(B31,' + twGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', twUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B30+B11', 'actual', '¥#,##0', '式: 原価(B30)+うわのせ(B11)'],
    ['実質粗利額', '=B14-B14*B28-B30', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B15/B14,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B16<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B30/(1-B10-B28),-1),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8480'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B30/(1-B10-B28),-2),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8500'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B30/(1-B31-B28),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B30/(1-B31-B28),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9', 'auto', '0%', 'モール+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '本の価格×送料率'],
    ['原価(円)', '=ROUND((B4+B29)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B30', 'B34', 'B35', 'B36', 'B37', 'B38'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 中国
  // 入力 B4-B12 / 微調整 B15-B18 / おすすめ B21-B22 / 原価帯売値 B25-B26
  // 自動計算: 手数料B29 送料B30 原価B31 適用B32 / 帯 B35-B39
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
    ['目標粗利率', '=IFERROR(B32,' + cnGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', cnUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B31+B12', 'actual', '¥#,##0', '式: 原価(B31)+うわのせ(B12)'],
    ['実質粗利額', '=B15-B15*B29-B31', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B16/B15,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B17<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B11-B29),-1),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8480'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B11-B29),-2),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8500'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B32-B29),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B32-B29),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9+B10', 'auto', '0%', 'モール+カード+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '商品価格×送料率'],
    ['原価(円)', '=ROUND((B4+B30)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B31', 'B35', 'B36', 'B37', 'B38', 'B39'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 韓国
  // 入力 B4-B12 / 微調整 B15-B18 / おすすめ B21-B22 / 原価帯売値 B25-B26
  // 自動計算: 手数料B29 原価B30 適用B31 / 帯 B34-B38
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
    ['目標粗利率', '=IFERROR(B31,' + krGoal + ')', 'input', '0%', '通常は原価帯自動に連動。今回だけ変えたい時は数値で上書き'],
    ['実際のうわのせ', krUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B30+B12', 'actual', '¥#,##0', '式: 原価(B30)+うわのせ(B12)'],
    ['実質粗利額', '=B15-B15*B29-B30', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B16/B15,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B17<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B30/(1-B11-B29),-1),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8480'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B30/(1-B11-B29),-2),0)', 'result', '¥#,##0', '目標粗利率で計算 例:8500'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B30/(1-B31-B29),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B30/(1-B31-B29),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で自動'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B9+B10', 'auto', '0%', 'モール+消費税'],
    ['原価(円)', '=ROUND(B4*B6+B5*B8,0)', 'auto', '¥#,##0', '価格×レート+重さ×係数'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B30', 'B34', 'B35', 'B36', 'B37', 'B38'), 'auto', '0%', '一番下の帯設定×原価から自動'],
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
    bandPrice: '#FCE4EC',
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
    if (type === 'section' || type === 'sectionPink') {
      sh.getRange(row, 1, 1, 3).merge();
      const sec = sh.getRange(row, 1).setValue(label).setFontWeight('bold').setFontSize(13);
      if (type === 'sectionPink') {
        sec.setFontColor('#C2185B').setBackground('#FCE4EC');
      } else {
        sec.setFontColor('#3C4043');
      }
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
    if (type === 'result' && label.indexOf('おすすめ') >= 0) {
      b.setFontSize(16).setFontWeight('bold');
    }
    if (type === 'bandPrice') {
      b.setFontSize(16).setFontWeight('bold');
    }
    if (type === 'actual' && (label === '実際売値' || label === '実質粗利額')) {
      b.setFontSize(18).setFontWeight('bold');
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
