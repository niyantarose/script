/**
 * 仕入れ計算機 v2（中国・韓国・台湾）
 *
 * ・原価帯別目標粗利率を自動適用（台湾・中国・韓国）
 * ・おすすめ（目標から）と微調整（実際）を別セクション
 * ・韓国: 縦×横×厚(cm)×0.45で重さ自動、未入力時はg手入力
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

/** スプレッドシート用 IFS 式（{cost} を原価セルに置換） */
function 計算機_原価帯粗利率式_(costA1) {
  return '=IFS(' + costA1 + '<=2000,0.35,' + costA1 + '<=6000,0.25,' + costA1 + '<=12000,0.2,' + costA1 + '<=20000,0.1,TRUE,0.08)';
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

/** 原価帯参考テーブル行 */
function 計算機_原価帯参考行_() {
  return [
    ['原価帯別目標粗利率（参考）', '', 'refTitle', '', '↑原価から自動で「適用目標粗利率」が決まります'],
    ['原価帯', '目標粗利率', 'refHead', '@', ''],
    ['〜 2,000円まで', 0.35, 'refRow', '0%', ''],
    ['〜 6,000円まで', 0.25, 'refRow', '0%', ''],
    ['〜 12,000円まで', 0.2, 'refRow', '0%', ''],
    ['〜 20,000円まで', 0.1, 'refRow', '0%', ''],
    ['〜 30,000円まで', 0.08, 'refRow', '0%', ''],
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

  const 台湾 = [
    ['台湾 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['本の価格(元)', 820, 'input', '0.0', '元'],
    ['使用レート(円/元)', '=IFERROR(B6,5.09)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:TWDJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.35, 'input', '0%', '送料(元)=本の価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9', 'auto', '0%', 'モール+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '本の価格×送料率'],
    ['原価(円)', '=ROUND((B4+B13)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(自動)', 計算機_原価帯粗利率式_('B14'), 'auto', '0%', '原価帯で自動決定'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯参考行_()).concat([
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B14/(1-B15-B12),0)', 'result', '¥#,##0', '式: 原価÷(1-適用目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B26,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B26,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B27-B14', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B28-B14', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['おすすめ時粗利額(10円丸め)', '=B27-B27*B12-B14', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['おすすめ時粗利額(100円丸め)', '=B28-B28*B12-B14', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', twUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B14+B35', 'actual', '¥#,##0', '式: 原価(B14)+うわのせ(B35)'],
    ['実質粗利額', '=B36-B36*B12-B14', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B37)/B36', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B38<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ]);

  const 中国 = [
    ['中国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['商品価格(元)', 37.4, 'input', '0.0', '元'],
    ['使用レート(円/元)', '=IFERROR(B6,23.82)', 'input', '0.000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:CNYJPY"),"取得不可")', 'auto', '0.000', '自動取得'],
    ['送料率', 0.3, 'input', '0%', '送料(元)=商品価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['カード手数料', 0.03, 'input', '0%', '決済手数料'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9+B10', 'auto', '0%', 'モール+カード+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '商品価格×送料率'],
    ['原価(円)', '=ROUND((B4+B14)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(自動)', 計算機_原価帯粗利率式_('B15'), 'auto', '0%', '原価帯で自動決定'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯参考行_()).concat([
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B15/(1-B16-B13),0)', 'result', '¥#,##0', '式: 原価÷(1-適用目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B27,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B27,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B28-B15', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B29-B15', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['おすすめ時粗利額(10円丸め)', '=B28-B28*B13-B15', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['おすすめ時粗利額(100円丸め)', '=B29-B29*B13-B15', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', cnUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B15+B36', 'actual', '¥#,##0', '式: 原価(B15)+うわのせ(B36)'],
    ['実質粗利額', '=B37-B37*B13-B15', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B38)/B37', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B39<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ]);

  const 韓国 = [
    ['韓国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色）──', '', 'section'],
    ['本の価格(ウォン)', 180900, 'input', '#,##0', 'ウォン'],
    ['縦(cm)', '', 'input', '0.0', '任意: 寸法から重さ自動'],
    ['横(cm)', '', 'input', '0.0', '任意'],
    ['厚さ(cm)', '', 'input', '0.0', '任意'],
    ['重さ(g)手入力', 300, 'input', '#,##0', '寸法未入力時はこちらを使用'],
    ['使用レート(円/ウォン)', '=IFERROR(B10,0.105)', 'input', '0.0000', '通常は自動。盛る時は数値で上書き'],
    ['市場レート(自動)', '=IFERROR(GOOGLEFINANCE("CURRENCY:KRWJPY"),"取得不可")', 'auto', '0.0000', '自動取得'],
    ['重量係数(送料)', 0.6, 'input', '0.0', '送料(円)=重さ×係数'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['使用する重さ(g)', '=IF(AND(B5>0,B6>0,B7>0),ROUND(B5*B6*B7*0.45,0),B8)', 'auto', '#,##0', '縦×横×厚×0.45 または手入力'],
    ['手数料合計', '=B12+B13', 'auto', '0%', 'モール+消費税'],
    ['原価(円)', '=ROUND(B4*B9+B16*B11,0)', 'auto', '¥#,##0', '価格×レート+重さ×係数'],
    ['適用目標粗利率(自動)', 計算機_原価帯粗利率式_('B18'), 'auto', '0%', '原価帯で自動決定'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯参考行_()).concat([
    ['', '', 'blank'],
    ['── おすすめ（目標粗利率から・緑）──', '', 'section'],
    ['理論売値', '=ROUND(B18/(1-B19-B17),0)', 'result', '¥#,##0', '式: 原価÷(1-目標粗利率-手数料合計)'],
    ['おすすめ売値(10円丸め)', '=ROUND(B30,-1)', 'result', '¥#,##0', '理論売値を10円単位に丸め'],
    ['おすすめ売値(100円丸め)', '=ROUND(B30,-2)', 'result', '¥#,##0', '理論売値を100円単位に丸め'],
    ['推奨うわのせ(10円丸め)', '=B31-B18', 'result', '¥#,##0', '式: おすすめ売値(10円)-原価'],
    ['推奨うわのせ(100円丸め)', '=B32-B18', 'result', '¥#,##0', '式: おすすめ売値(100円)-原価'],
    ['おすすめ時粗利額(10円丸め)', '=B31-B31*B17-B18', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['おすすめ時粗利額(100円丸め)', '=B32-B32*B17-B18', 'result', '¥#,##0', 'この売値なら手数料後にこれだけ残る'],
    ['', '', 'blank'],
    ['── 微調整（実際・青）──', '', 'section'],
    ['実際のうわのせ', krUwa, 'actual', '¥#,##0', '↑推奨をコピーして手で調整'],
    ['実際売値', '=B18+B39', 'actual', '¥#,##0', '式: 原価(B18)+うわのせ(B39)'],
    ['実質粗利額', '=B40-B40*B17-B18', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=(B41)/B40', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B42<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
  ]);

  計算機_buildSheet_(ss, '台湾計算機', 台湾);
  計算機_buildSheet_(ss, '韓国計算機', 韓国);
  計算機_buildSheet_(ss, '中国計算機', 中国);

  ss.toast('台湾・韓国・中国 計算機 v2 を再作成しました（原価帯自動）');
}

function 計算機_setNote_(range, note) {
  if (!note) return;
  const text = String(note);
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
  let refTableStart = null;
  let refTableEnd = null;
  let refRowToggle = false;

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
    if (type === 'refTitle' || type === 'refHead' || type === 'refRow') {
      if (refTableStart === null) refTableStart = row;
      refTableEnd = row;
      const a = sh.getRange(row, 1);
      const b = sh.getRange(row, 2);
      a.setValue(label);
      if (type === 'refTitle') {
        sh.getRange(row, 1, 1, 2).merge();
        a.setFontWeight('bold').setFontSize(11).setFontColor('#202124')
          .setBackground('#DADCE0').setHorizontalAlignment('center');
        if (note) 計算機_setNote_(sh.getRange(row, 3), note);
        return;
      }
      if (type === 'refHead') {
        a.setFontWeight('bold').setFontSize(10).setFontColor('#3C4043').setBackground('#ECEFF1');
        b.setValue(val).setFontWeight('bold').setFontSize(10).setFontColor('#3C4043')
          .setBackground('#ECEFF1').setHorizontalAlignment('center');
        return;
      }
      refRowToggle = !refRowToggle;
      const rowBg = refRowToggle ? '#FFFFFF' : '#F8F9FA';
      a.setFontSize(11).setFontColor('#202124').setBackground(rowBg);
      b.setValue(val);
      if (fmt) b.setNumberFormat(fmt);
      b.setFontSize(12).setFontWeight('bold').setFontColor('#1A73E8')
        .setBackground(rowBg).setHorizontalAlignment('center');
      return;
    }
    const a = sh.getRange(row, 1);
    const b = sh.getRange(row, 2);
    a.setValue(label);
    if (typeof val === 'string' && val.charAt(0) === '=') b.setFormula(val);
    else if (val === '' && type === 'input') b.setValue('');
    else b.setValue(val);
    if (fmt) b.setNumberFormat(fmt);
    if (BG[type]) {
      a.setBackground(BG[type]);
      b.setBackground(BG[type]);
    }
    if (type === 'result' && (label.indexOf('おすすめ') >= 0 || label.indexOf('理論') >= 0)) {
      b.setFontSize(13).setFontWeight('bold');
    }
    if (type === 'actual' && label === '実際売値') {
      b.setFontSize(13).setFontWeight('bold');
    }
    if (note) 計算機_setNote_(sh.getRange(row, 3), note);
    if (label === '実質粗利率' && type === 'actual') profitA1 = b.getA1Notation();
  });

  if (refTableStart && refTableEnd) {
    const tbl = sh.getRange(refTableStart, 1, refTableEnd - refTableStart + 1, 2);
    tbl.setBorder(true, true, true, true, true, true, '#BDC1C6', SpreadsheetApp.BorderStyle.SOLID);
  }

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
