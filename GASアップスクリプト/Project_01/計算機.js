/**
 * 仕入れ計算機 v2（中国・韓国・台湾）
 *
 * レイアウト: 入力 → 微調整 → おすすめ(10円/100円) → 自動価格計算結果(原価帯) → 【別設定】原価帯(最下)
 *
 * 2系統は独立している（連動しない）:
 *   - 入力の「目標粗利率」= 自分で決める率。緑「おすすめ売値」だけに使う手動の試算用
 *   - ピンク「原価帯売値」= 最下部の原価帯ルールから全自動。入力の目標粗利率は参照しない
 * 入力の「実際のうわのせ」の真下に、その上乗せでの売値をその場で表示する。
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

/** 超軽量動作確認（止まっている時はまずこれ） */
function 計算機_動作確認() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('計算機スクリプトは動いています', '計算機', 5);
  Logger.log('計算機_動作確認 OK ' + new Date());
}

function 計算機_きれいなシートを作成() {
  try {
    計算機_きれいなシートを作成_();
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    Logger.log('計算機シート作成エラー: ' + msg);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('エラー: ' + msg, '計算機', 15);
    } catch (ignore) {}
    try {
      SpreadsheetApp.getUi().alert('計算機シート作成エラー', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (ignore2) {}
  }
}

function 計算機_きれいなシートを作成_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('開始: 台湾→韓国→中国', '計算機', 8);
  SpreadsheetApp.flush();

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
  // 入力 B4-B12（B12=この上乗せでの売値） / 微調整 B15-B18 / おすすめ B21-B22 / 原価帯売値 B25-B26
  // 自動計算: 手数料B29 送料B30 原価B31 適用B32 / 帯 B35-B39
  const 台湾 = [
    ['台湾 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['本の価格(元)', 820, 'input', '0.0', '元'],
    ['使用レート(円/元)', 5.09, 'input', '0.000', '市場レートを見て必要なら上書き'],
    ['市場レート(目安)', 5.09, 'auto', '0.000', '目安。正確なレートは別途確認'],
    ['送料率', 0.35, 'input', '0%', '送料(元)=本の価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', twGoal, 'input', '0%', '自分で決める率。原価帯とは連動しません（緑「おすすめ売値」の計算に使用）'],
    ['実際のうわのせ', twUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['→ この上乗せでの売値', '=B31+B11', 'actual', '¥#,##0', 'いま入れているうわのせでの売値（原価+うわのせ）'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B31+B11', 'actual', '¥#,##0', '式: 原価(B31)+うわのせ(B11)'],
    ['実質粗利額', '=B15-B15*B29-B31', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B16/B15,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B17<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B10-B29),-1),0)', 'result', '¥#,##0', '入力の目標粗利率(B10)で計算'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B10-B29),-2),0)', 'result', '¥#,##0', '入力の目標粗利率(B10)で計算'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B32-B29),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B32-B29),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9', 'auto', '0%', 'モール+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '本の価格×送料率'],
    ['原価(円)', '=ROUND((B4+B30)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B31', 'B35', 'B36', 'B37', 'B38', 'B39'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 中国
  // 入力 B4-B13（B13=この上乗せでの売値） / 微調整 B16-B19 / おすすめ B22-B23 / 原価帯売値 B26-B27
  // 自動計算: 手数料B30 送料B31 原価B32 適用B33 / 帯 B36-B40
  const 中国 = [
    ['中国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['商品価格(元)', 37.4, 'input', '0.0', '元'],
    ['使用レート(円/元)', 23.82, 'input', '0.000', '市場レートを見て必要なら上書き'],
    ['市場レート(目安)', 23.82, 'auto', '0.000', '目安。正確なレートは別途確認'],
    ['送料率', 0.3, 'input', '0%', '送料(元)=商品価格×送料率'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['カード手数料', 0.03, 'input', '0%', '決済手数料'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', cnGoal, 'input', '0%', '自分で決める率。原価帯とは連動しません（緑「おすすめ売値」の計算に使用）'],
    ['実際のうわのせ', cnUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['→ この上乗せでの売値', '=B32+B12', 'actual', '¥#,##0', 'いま入れているうわのせでの売値（原価+うわのせ）'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B32+B12', 'actual', '¥#,##0', '式: 原価(B32)+うわのせ(B12)'],
    ['実質粗利額', '=B16-B16*B30-B32', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B17/B16,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B18<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B32/(1-B11-B30),-1),0)', 'result', '¥#,##0', '入力の目標粗利率(B11)で計算'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B32/(1-B11-B30),-2),0)', 'result', '¥#,##0', '入力の目標粗利率(B11)で計算'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B32/(1-B33-B30),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B32/(1-B33-B30),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B8+B9+B10', 'auto', '0%', 'モール+カード+消費税'],
    ['送料(元)', '=B4*B7', 'auto', '0.0', '商品価格×送料率'],
    ['原価(円)', '=ROUND((B4+B31)*B5,0)', 'auto', '¥#,##0', '(価格+送料)×レート'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B32', 'B36', 'B37', 'B38', 'B39', 'B40'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  // 韓国
  // 入力 B4-B13（B13=この上乗せでの売値） / 微調整 B16-B19 / おすすめ B22-B23 / 原価帯売値 B26-B27
  // 自動計算: 手数料B30 原価B31 適用B32 / 帯 B35-B39
  const 韓国 = [
    ['韓国 仕入れ計算機 v2', '', 'title'],
    ['', '', 'blank'],
    ['── 入力（黄色・毎回ここ）──', '', 'section'],
    ['本の価格(ウォン)', 180900, 'input', '#,##0', 'ウォン'],
    ['重さ(g)', 300, 'input', '#,##0', 'グラムで重さを入力'],
    ['使用レート(円/ウォン)', 0.105, 'input', '0.0000', '市場レートを見て必要なら上書き'],
    ['市場レート(目安)', 0.105, 'auto', '0.0000', '目安。正確なレートは別途確認'],
    ['重量係数(送料)', 0.6, 'input', '0.0', '送料(円)=重さ×係数'],
    ['モール手数料', 0.09, 'input', '0%', 'Yahoo等'],
    ['消費税', 0.1, 'input', '0%', '0.1=10%'],
    ['目標粗利率', krGoal, 'input', '0%', '自分で決める率。原価帯とは連動しません（緑「おすすめ売値」の計算に使用）'],
    ['実際のうわのせ', krUwa, 'input', '¥#,##0', '手で調整。初期値はおすすめ(10円丸め)基準'],
    ['→ この上乗せでの売値', '=B31+B12', 'actual', '¥#,##0', 'いま入れているうわのせでの売値（原価+うわのせ）'],
    ['', '', 'blank'],
    ['── 微調整結果（青）──', '', 'section'],
    ['実際売値', '=B31+B12', 'actual', '¥#,##0', '式: 原価(B31)+うわのせ(B12)'],
    ['実質粗利額', '=B16-B16*B30-B31', 'actual', '¥#,##0', '式: 売値-売値×手数料合計-原価'],
    ['実質粗利率', '=IFERROR(B17/B16,0)', 'actual', '0%', '20%以上が目安'],
    ['判定', '=IF(B18<0.2,"⚠ 20%割れ 注意","OK")', 'actual', '@', '微調整後の粗利率で判定'],
    ['', '', 'blank'],
    ['── おすすめ（緑）──', '', 'section'],
    ['おすすめ売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B11-B30),-1),0)', 'result', '¥#,##0', '入力の目標粗利率(B11)で計算'],
    ['おすすめ売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B11-B30),-2),0)', 'result', '¥#,##0', '入力の目標粗利率(B11)で計算'],
    ['', '', 'blank'],
    ['── 自動価格計算結果（原価帯）──', '', 'sectionPink'],
    ['原価帯売値(10円丸め)', '=IFERROR(ROUND(B31/(1-B32-B30),-1),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['原価帯売値(100円丸め)', '=IFERROR(ROUND(B31/(1-B32-B30),-2),0)', 'bandPrice', '¥#,##0', '原価帯の適用率で全自動（入力の目標粗利率は使いません）'],
    ['', '', 'blank'],
    ['── 自動計算 ──', '', 'section'],
    ['手数料合計', '=B9+B10', 'auto', '0%', 'モール+消費税'],
    ['原価(円)', '=ROUND(B4*B6+B5*B8,0)', 'auto', '¥#,##0', '価格×レート+重さ×係数'],
    ['適用目標粗利率(原価帯から自動)', 計算機_原価帯粗利率式_('B31', 'B35', 'B36', 'B37', 'B38', 'B39'), 'auto', '0%', '一番下の帯設定×原価から自動'],
    ['', '', 'blank'],
  ].concat(計算機_原価帯設定行_());

  計算機_buildSheetFast_(ss, '台湾計算機', 台湾);
  ss.toast('台湾OK → 韓国…', '計算機', 5);
  SpreadsheetApp.flush();

  計算機_buildSheetFast_(ss, '韓国計算機', 韓国);
  ss.toast('韓国OK → 中国…', '計算機', 5);
  SpreadsheetApp.flush();

  計算機_buildSheetFast_(ss, '中国計算機', 中国);

  const done = ss.getSheetByName('台湾計算機');
  if (done) ss.setActiveSheet(done);
  ss.toast('完了：台湾計算機 / 韓国計算機 / 中国計算機', '計算機', 8);
}

/**
 * 既存シートを clear して一括書き込み（行ごとの API 連打を避ける）
 */
function 計算機_buildSheetFast_(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clear();
    sh.setConditionalFormatRules([]);
  }

  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 400);

  const BG = {
    input: '#FFF2CC',
    band: '#FCE8B2',
    auto: '#F1EFE8',
    result: '#D9EAD3',
    actual: '#D0E0F0',
    bandPrice: '#FCE4EC',
  };

  const n = rows.length;
  const values = [];
  const bgA = [];
  const bgB = [];
  const bgC = [];
  const fmtB = [];
  const sizeA = [];
  const sizeB = [];
  const weightA = [];
  const weightB = [];
  const colorA = [];
  const formulaRows = [];
  let profitA1 = null;

  for (let i = 0; i < n; i++) {
    const r = rows[i];
    const type = r[2] || 'blank';
    const label = (type === 'blank') ? '' : (r[0] || '');
    const noteRaw = r[4] ? String(r[4]) : '';
    const note = noteRaw.charAt(0) === '=' ? "'" + noteRaw : noteRaw;
    const val = r[1];
    const isFormula = typeof val === 'string' && val.charAt(0) === '=';
    const isHeader = (type === 'title' || type === 'section' || type === 'sectionPink' || type === 'blank');

    values.push([label, isHeader || isFormula ? '' : (val === undefined || val === null ? '' : val), note]);

    let ba = null;
    let bb = null;
    let bc = null;
    let fa = 13;
    let fb = 14;
    let wa = 'normal';
    let wb = 'normal';
    let ca = '#000000';

    if (type === 'title') {
      fa = 18; wa = 'bold';
    } else if (type === 'section') {
      fa = 13; wa = 'bold'; ca = '#3C4043';
    } else if (type === 'sectionPink') {
      fa = 13; wa = 'bold'; ca = '#C2185B';
      ba = '#FCE4EC'; bb = '#FCE4EC'; bc = '#FCE4EC';
    } else if (BG[type]) {
      ba = BG[type];
      bb = BG[type];
      if (type === 'result' && String(r[0]).indexOf('おすすめ') >= 0) {
        fb = 16; wb = 'bold';
      }
      if (type === 'bandPrice') {
        fb = 16; wb = 'bold';
      }
      if (type === 'actual' && (r[0] === '実際売値' || r[0] === '実質粗利額')) {
        fb = 18; wb = 'bold';
      } else if (type === 'actual' && String(r[0]).indexOf('この上乗せでの売値') >= 0) {
        fb = 16; wb = 'bold';
      }
      if (type === 'band') wb = 'bold';
    }

    bgA.push([ba]);
    bgB.push([bb]);
    bgC.push([bc]);
    fmtB.push([(!isHeader && r[3]) ? r[3] : '@']);
    sizeA.push([fa]);
    sizeB.push([fb]);
    weightA.push([wa]);
    weightB.push([wb]);
    colorA.push([ca]);

    if (isFormula) formulaRows.push({ row: i + 1, formula: val });
    if (r[0] === '実質粗利率' && type === 'actual') profitA1 = 'B' + (i + 1);
  }

  const rng = sh.getRange(1, 1, n, 3);
  rng.setValues(values);
  sh.getRange(1, 1, n, 1).setBackgrounds(bgA).setFontSizes(sizeA).setFontWeights(weightA).setFontColors(colorA);
  sh.getRange(1, 2, n, 1).setBackgrounds(bgB).setFontSizes(sizeB).setFontWeights(weightB).setNumberFormats(fmtB);
  sh.getRange(1, 3, n, 1).setBackgrounds(bgC).setFontColor('#5F6368').setFontSize(11);

  for (let i = 0; i < formulaRows.length; i++) {
    sh.getRange(formulaRows[i].row, 2).setFormula(formulaRows[i].formula);
  }

  // タイトル・見出しだけ結合（少なめ）
  for (let i = 0; i < n; i++) {
    const type = rows[i][2];
    if (type === 'title' || type === 'section' || type === 'sectionPink') {
      sh.getRange(i + 1, 1, 1, 3).merge();
    }
  }

  if (profitA1) {
    const pr = sh.getRange(profitA1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.2).setBackground('#F4C7C3').setFontColor('#CC0000').setBold(true)
        .setRanges([pr]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.2).setBackground('#D0E0F0')
        .setRanges([pr]).build(),
    ]);
  }
}
