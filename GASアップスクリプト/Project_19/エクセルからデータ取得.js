// ====== 手動運用メニュー ======
function Excel同期メニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('手動運用')
    .addItem('表示：最終データ行へ移動', '最終データ行へ移動')
    .addSeparator()
    .addItem('発注リスト大邱：一意No採番', '大邱発注_一意Noを採番')
    .addItem('発注リスト大邱：グループ罫線', '大邱発注_グループ罫線')
    .addItem('発注リスト大邱：チェック＆残り数量＆消込色を設置', '大邱発注_チェックと残り数量を設置')
    .addItem('入荷数あり行：入荷日を補完', '入荷数あり行_入荷日を補完')
    .addItem('チェック行：入荷数を発注数量にする', 'チェック行_入荷数を発注数量にする')
    .addItem('発注リスト大邱：チェックを全部外す', '大邱発注_チェックを全部外す')
    .addItem('発注リスト大邱：フィルタを最新化(全表示)', '大邱発注_フィルタを最新化')
    .addItem('発注リスト大邱：同一コードの重量を補完', '大邱発注_同一コードの重量を補完')
    .addItem('発注リスト大邱：Q/S/W/Xを再計算', '大邱発注_WXを全体再計算')
    .addItem('発注リスト大邱：チェック行をEMS大邱へ送る', '大邱発注_チェック行をEMS大邱へ送る')
    .addItem('大邱：発注No枝番をEMS大邱まで一括修正', '大邱_発注EMS枝番を一括修正')
    .addItem('大邱データに発注/EMSリストを合わせる', '大邱データに発注とEMSを合わせる')
    .addItem('EMS大邱：購入Noを初回補完（発注大邱から）', 'EMS大邱_購入No初回補完')
    .addItem('EMS大邱：N/Q/R/Sを再計算', 'EMS大邱_QRSを全体再計算')
    .addItem('EMS大邱：EMS番号ごとの罫線を更新', 'EMS大邱_EMS番号ごとに罫線を更新')
    .addItem('EMS大邱：シール発行済みに一括初期化(既存行)', 'EMS大邱_シール発行済みに一括初期化')
    .addItem('EMS大邱：入荷日を発注リストから補完', 'EMS大邱_入荷日を発注リストから補完')
    .addItem('発注リスト大邱データ → 発注', '大邱_発注へ転送')
    .addItem('発注：背景色(行ハイライト)を大邱に合わせる', '大邱_発注の背景色を大邱に合わせる')
    .addItem('EMS大邱作業データ → EMSリスト', '大邱_EMSリストへ転送')
    .addItem('EMS番号：大邱/EMSリストを正規化', 'EMS番号_大邱とEMSリストを正規化')
    .addItem('EMSリスト：大邱と箱数を照合', 'EMSリスト_大邱と箱数を照合')
    .addItem('EMSリスト：大邱から不足行を確認して復元', 'EMSリスト_大邱から不足行を確認して復元')
    .addItem('EMSリスト：発送日基準に並べ替え', 'EMSリスト_発送日基準に並べ替え')
    .addItem('EMSリスト：品目を生データ化（関数を外す）', 'EMSリスト_品目を生データ化')
    .addItem('購入No補完（EMSリスト）', 'EMSリスト_購入No自動補完')
    .addItem('発注：EMS発送数を再計算', '発注_EMS発送数数式を一括修正')
    .addItem('発送数：自動再計算トリガーを設置', '発送数自動再計算トリガーを設置')
    .addSeparator()
    .addItem('EMSカレンダーシートを更新', 'buildEmsCalendarSheet')
    .addItem('EMSカレンダー：当日ハイライトを更新', 'refreshEmsCalendarTodayHighlight_')
    .addItem('GoogleカレンダーへEMSを反映', 'syncEmsCalendar')
    .addItem('EMSリスト：重複行を確認して削除', 'EMSリスト_重複行を確認して削除')
    .addItem('EMSリスト：下に混ざった古い行を確認して削除', 'EMSリスト_下に混ざった古い行を確認して削除')
    .addItem('商品マスタ：重複/商品コードなしを整理', '商品マスタ_商品コードを一意化')
    .addSeparator()
    .addItem('商品マスタ：発注大邱データから更新', '商品マスタ_発注大邱から更新')
    .addItem('商品マスタ：新リスト(取込タブ)をマージ', '商品マスタ_新リストをマージ')
    .addItem('商品マスタ：品目を英語に統一', '商品マスタ_品目を英語に統一')
    .addItem('商品マスタ：除外コードシートを作成/更新', '商品マスタ_除外コードシートを作成')
    .addItem('商品マスタ：人名ラベル行を削除(英数字なし)', '商品マスタ_人名ラベル行を削除')
    .addSeparator()
    .addItem('旧Excel同期トリガーを削除', '旧Excel同期トリガーを削除')
    .addToUi();
}

function 旧Excel同期トリガーを削除() {
  const deleted = disableLegacyExcelSyncTriggers_();
  SpreadsheetApp.getActive().toast(`旧Excel同期トリガーを削除: ${deleted}件`, '手動運用', 5);
}

function disableLegacyExcelSyncTriggers_() {
  const legacy = { syncAll: true, syncDifferences: true, importSharePointExcel: true };
  let deleted = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (legacy[t.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });
  return deleted;
}

function setupTriggers() {
  const deleted = disableLegacyExcelSyncTriggers_();
  Logger.log(`Excel同期は廃止済みです。旧トリガーを ${deleted}件 削除しました。`);
}

function syncAll() {
  const deleted = disableLegacyExcelSyncTriggers_();
  Logger.log(`syncAllは廃止済みです。旧トリガーを ${deleted}件 削除しました。`);
}

// ---- 正規化ヘルパー ----
function normCode_(v) {
  return String(v || '')
    .normalize('NFKC')
    .replace(/[‐‑‒–—―−]/g, '-')
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .trim()
    .toUpperCase()
    .replace(/_/g, '-')
    .replace(/-+/g, '-');
}
function normTrack_(v) {
  return String(v || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '' || String(v).trim() === '　';
}

// ====== EMSリスト → Googleカレンダー同期 ======
const EMS_CAL_CFG = {
  CALENDAR_NAME: 'EMS追跡',
  MARKER: '[EMS-SYNC]',  // スクリプトが作ったイベントの目印
  COLORS: {
    '未着':       CalendarApp.EventColor.BLUE,
    '輸送中':     CalendarApp.EventColor.ORANGE,
    '在庫反映済み': CalendarApp.EventColor.GREEN
  }
};

function syncEmsCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('EMSリスト');
  if (!sh) return;

  const vals = sh.getDataRange().getValues();
  let header = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') { header = i; break; }
  }
  if (header < 0) return;

  // EMS番号＋Box単位でグループ化
  const groups = {};
  for (let i = header + 1; i < vals.length; i++) {
    const r = vals[i];
    const track = normTrack_(r[12]);              // M列 EMS番号
    if (!track) continue;
    const ship = _emsDateOnly_(r[2]);             // C列 発送日
    const arrive = _emsDateOnly_(r[4]);           // E列 到着日
    if (!ship) continue;                          // 発送日が無い行はスキップ
    const key = track + '|' + String(r[13]);      // Box No.込みでキー化
    if (!groups[key]) {
      groups[key] = {
        track: track,
        box: r[13],
        ship: ship,
        arrive: (arrive instanceof Date) ? arrive : null,
        status: String(r[6] || '').trim(),        // G列 ステータス
        items: [], qty: 0
      };
    } else {
      if (ship < groups[key].ship) groups[key].ship = ship;
      if (arrive && (!groups[key].arrive || arrive > groups[key].arrive)) groups[key].arrive = arrive;
    }
    groups[key].items.push(`${r[8]} ×${r[9]}（${r[10]}）`); // 商品コード×数量（品目）
    groups[key].qty += Number(r[9]) || 0;
    // 1行でも「未着」があれば箱全体を未着扱い
    if (String(r[6]).trim() === '未着') groups[key].status = '未着';
  }

  // 専用カレンダーを取得（なければ作る）
  let cal = CalendarApp.getCalendarsByName(EMS_CAL_CFG.CALENDAR_NAME)[0];
  if (!cal) cal = CalendarApp.createCalendar(EMS_CAL_CFG.CALENDAR_NAME);

  // スクリプト製の既存イベントを全削除して作り直す（常にシートが正）
  cal.getEvents(new Date('2026-01-01'), new Date('2031-01-01')).forEach(ev => {
    if (ev.getDescription().indexOf(EMS_CAL_CFG.MARKER) >= 0) ev.deleteEvent();
  });

  // 箱ごとに「発送日〜到着日」の帯イベントを作成
  // 箱ごとに「発送日〜到着日」の帯イベントを作成
  let count = 0;
  const today = _emsTodayTokyo_();   // ★追加
  Object.keys(groups).forEach(key => {
    const g = groups[key];
    const start = g.ship;
    const endBase = g.arrive || g.ship;
    const end = new Date(endBase.getFullYear(), endBase.getMonth(), endBase.getDate() + 1);

    const title = `【${g.status || '状態不明'}】${g.track} Box${g.box}（${g.items.length}品目 ${g.qty}点）`;
    const desc = EMS_CAL_CFG.MARKER + '\n' +
      '追跡: https://trackings.post.japanpost.jp/services/srv/search/input\n' +
      g.items.join('\n');

    const ev = cal.createAllDayEvent(title, start, end, { description: desc });
    const status = String(g.status || '').trim();
    const isInactive = status === '在庫反映済み' || (g.arrive instanceof Date && g.arrive < today);
    const color = isInactive                               // ★到着済み/在庫反映済みはグレー
      ? CalendarApp.EventColor.GRAY
      : EMS_CAL_CFG.COLORS[g.status];
    if (color) ev.setColor(color);
    count++;
  });

  Logger.log(`EMSカレンダー同期完了: ${count}箱`);
}

// ---- EMSリストを箱単位にまとめる（カレンダー系の共通処理）----
function getEmsBoxGroups_(ss) {
  const sh = ss.getSheetByName('EMSリスト');
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  const displayVals = sh.getDataRange().getDisplayValues();
  let header = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') { header = i; break; }
  }
  if (header < 0) return {};

  // M:N列（EMS番号/Box No.）の背景色＝EMSリストの箱グループ色を帯にも使う
  const emsBgs = sh.getRange(1, 13, vals.length, 2).getBackgrounds();

  const groups = {};
  const trackArrivals = {};
  for (let i = header + 1; i < vals.length; i++) {
    const r = vals[i];
    const d = displayVals[i];
    const track = normTrack_(r[12]);
    if (!track) continue;
    const ship = _emsDateOnly_(r[2]) || _emsDateOnly_(d[2]);
    const arrive = _emsDateOnly_(r[4]) || _emsDateOnly_(d[4]);
    if (arrive && (!trackArrivals[track] || arrive > trackArrivals[track])) {
      trackArrivals[track] = arrive;
    }
    if (!ship) continue;
    const box = String(r[13] || d[13] || '1').trim();
    const key = track + '|' + box;
    const rowColor = _emsCalendarFirstUsefulBg_(emsBgs[i][1], emsBgs[i][0]);
    if (!groups[key]) {
      groups[key] = {
        track: track, box: box, ship: ship,
        arrive: (arrive instanceof Date) ? arrive : null,
        status: String(r[6] || '').trim(),
        color: rowColor,
        items: [], qty: 0
      };
    } else {
      if (ship < groups[key].ship) groups[key].ship = ship;
      if (arrive && (!groups[key].arrive || arrive > groups[key].arrive)) groups[key].arrive = arrive;
      if (_emsCalendarHasUsefulBg_(rowColor) && !_emsCalendarHasUsefulBg_(groups[key].color)) {
        groups[key].color = rowColor;
      }
    }
    groups[key].items.push(`${r[8]} ×${r[9]}（${r[10]}）`);
    groups[key].qty += Number(r[9]) || 0;
    if (String(r[6]).trim() === '未着') groups[key].status = '未着';
  }
  Object.keys(groups).forEach(key => {
    const g = groups[key];
    if (!g.arrive && trackArrivals[g.track]) g.arrive = trackArrivals[g.track];
  });
  return groups;
}

// 表示する期間（今日を基準に前後何ヶ月か）
const EMS_CAL_VIEW = { MONTHS_BACK: 1, MONTHS_FWD: 1 };
const EMS_CAL_TODAY = {
  BG: '#fde293',
  FONT: '#3c4043',
  BORDER: '#f9ab00',
  MARKS_KEY: 'EMS_CAL_TODAY_MARKS'
};

function buildEmsCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groups = getEmsBoxGroups_(ss);
  let boxes = Object.keys(groups).map(k => groups[k]);
  if (boxes.length === 0) { ss.toast('EMSデータがありません'); return; }

  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const fmtMd = d => Utilities.formatDate(d, 'Asia/Tokyo', 'M/d');
  const STYLE = {
    '未着':         { bg: '#dadce0', font: '#3c4043' },
    '輸送中':       { bg: '#f29900', font: '#ffffff' },
    '在庫反映済み': { bg: '#b7dec7', font: '#3c4043' },
    'default':      { bg: '#dadce0', font: '#3c4043' }
  };
  const PAST_STYLE = { bg: '#e8eaed', font: '#9aa0a6' };

  const today = _emsTodayTokyo_();

  // 表示ウィンドウ（先月初〜2ヶ月先の月末）
  const winStart = new Date(today.getFullYear(), today.getMonth() - EMS_CAL_VIEW.MONTHS_BACK, 1);
  const winEnd = new Date(today.getFullYear(), today.getMonth() + EMS_CAL_VIEW.MONTHS_FWD + 1, 0);

  // 箱の整形＋ウィンドウ外を除外
  boxes.forEach(b => {
    const status = String(b.status || '').trim();
    b.statusNorm = status;
    b.isOpenEnded = status === '未着' && !(b.arrive instanceof Date);
    b.end = b.arrive || (b.isOpenEnded ? winEnd : b.ship);
  });
  boxes = boxes.filter(b => b.end >= winStart && b.ship <= winEnd);

  boxes.forEach(b => {
    const status = b.statusNorm || String(b.status || '').trim();
    b.isPast = b.arrive instanceof Date && b.arrive < today;
    const isInactive = status === '在庫反映済み' || status === '到着済み' || b.isPast;
    const hasColor = _emsCalendarHasUsefulBg_(b.color);
    b.style = isInactive
      ? PAST_STYLE
      : (hasColor ? { bg: b.color, font: '#3c4043' } : STYLE['default']);
    const arriveText = b.arrive instanceof Date ? fmt(b.arrive) : (b.isOpenEnded ? '未定' : fmt(b.end));
    b.note = `【${b.status || '?'}】${b.track} Box${b.box}\n` +
             `発送: ${fmt(b.ship)} → 到着: ${arriveText}\n` +
             `${b.qty}点\n` + b.items.join('\n');
  });

  // レーン割当（重なる箱は別の段へ）
  boxes.sort((a, b) => a.ship - b.ship);
  const laneEnd = [];
  boxes.forEach(b => {
    let lane = laneEnd.findIndex(e => e < b.ship.getTime());
    if (lane < 0) { lane = laneEnd.length; laneEnd.push(0); }
    laneEnd[lane] = b.end.getTime();
    b.lane = lane;
  });

  // 帯の長さに応じたラベル（ステータスは常に先頭に表示）
  const bandLabel = (b, days) => {
    const arrive = b.arrive ? `${fmtMd(b.end)}着 ` : '';
    if (days >= 4) return `【${b.status || '?'}】${arrive}${b.track} B${b.box}（${b.qty}点）`;
    if (days >= 2) return `【${b.status || '?'}】${arrive}${b.track} B${b.box}`;
    return `【${b.status || '?'}】${arrive}${b.track}`;
  };

  let sh = ss.getSheetByName('EMSカレンダー');
  if (!sh) sh = ss.insertSheet('EMSカレンダー');
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearNotes();

  const youbi = ['日', '月', '火', '水', '木', '金', '土'];
  let row = 1;
  let cur = new Date(winStart);
  const allTodayMarks = [];

  while (cur <= winEnd) {
    const y = cur.getFullYear(), mo = cur.getMonth();
    const monthStart = new Date(y, mo, 1);
    const daysInMonth = new Date(y, mo + 1, 0).getDate();
    const monthEnd = new Date(y, mo, daysInMonth);
    const blockStart = row;
    const todayHighlights = [];

    sh.getRange(row, 1, 1, 7).merge().setValue(`'${y}年${mo + 1}月`)
      .setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#f1f3f4');
    row++;

    sh.getRange(row, 1, 1, 7).setValues([youbi])
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setBackground('#fafafa');
    sh.getRange(row, 1).setFontColor('#d93025');
    sh.getRange(row, 7).setFontColor('#1a73e8');
    row++;

    let weekStart = new Date(y, mo, 1);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay());

    while (weekStart <= monthEnd) {
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);

      // 日付番号行（他の月の日はグレー薄表示）
      const numRow = row;
      const numRange = sh.getRange(numRow, 1, 1, 7);
      numRange.clearContent().setNumberFormat('0');
      let todayCol = 0;
      for (let c = 0; c < 7; c++) {
        const d = new Date(weekStart); d.setDate(d.getDate() + c);
        const cell = sh.getRange(numRow, c + 1);
        if (d.getMonth() === mo) {
          cell.setValue(d.getDate());
          if (_emsSameDate_(d, today)) todayCol = c + 1;
        } else {
          cell.clearContent().setBackground('#f8f9fa');
        }
      }
      numRange
        .setNumberFormat('0')
        .setFontSize(14)
        .setHorizontalAlignment('right').setFontColor('#70757a');
      sh.setRowHeight(numRow, 26);
      row++;

      // この週×この月にかかる箱だけ描く（月またぎの二重描画を防止）
      const segs = [];
      boxes.forEach(b => {
        let s = b.ship > weekStart ? b.ship : weekStart;
        let e = b.end < weekEnd ? b.end : weekEnd;
        if (s < monthStart) s = monthStart;
        if (e > monthEnd) e = monthEnd;
        if (s > e) return;
        segs.push({ b: b, s: s, e: e });
      });

      const maxLane = segs.reduce((m, x) => Math.max(m, x.b.lane), -1);
      const laneRows = Math.max(maxLane + 1, 1);
      for (let L = 0; L < laneRows; L++) sh.setRowHeight(row + L, 32);  // ★32に
      if (todayCol) todayHighlights.push({ row: numRow, col: todayCol, rows: laneRows + 2 });

      // 1日ごとの外枠。日付行＋帯行＋余白行を曜日列ごとに囲む。
      for (let c = 1; c <= 7; c++) {
        sh.getRange(numRow, c, laneRows + 2, 1)
          .setBorder(true, true, true, true, null, null, '#d0d7de', SpreadsheetApp.BorderStyle.SOLID);
      }

      segs.forEach(x => {
        const c1 = Math.floor((x.s - weekStart) / 86400000) + 1;
        const c2 = Math.floor((x.e - weekStart) / 86400000) + 1;
        const days = c2 - c1 + 1;
        const range = sh.getRange(row + x.b.lane, c1, 1, days);
        if (days > 1) range.merge();
        range.setValue(bandLabel(x.b, days))
          .setBackground(x.b.style.bg).setFontColor(x.b.style.font)
          .setFontSize(12)
          .setVerticalAlignment('middle')
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
          .setNote(x.b.note);
      });

      row += laneRows;
      sh.setRowHeight(row, 8);
      row++;

      weekStart = new Date(weekEnd);
      weekStart.setDate(weekStart.getDate() + 1);
    }

    sh.getRange(blockStart, 1, row - blockStart, 7)
      .setBorder(true, true, true, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
    todayHighlights.forEach(h => _emsApplyTodayMark_(sh, h));
    allTodayMarks.push.apply(allTodayMarks, todayHighlights);

    row++;
    cur = new Date(y, mo + 1, 1);
  }

  for (let c = 1; c <= 7; c++) sh.setColumnWidth(c, 200);
  sh.setHiddenGridlines(true);
  _emsSaveTodayMarks_(allTodayMarks);

  // シート構築のあと、Googleカレンダー「EMS追跡」へも同期する（失敗してもシート更新は完了させる）
  let calMsg = '';
  try {
    syncEmsCalendar();
    calMsg = ' ＋ Googleカレンダー同期OK';
  } catch (err) {
    calMsg = ' （Googleカレンダー同期は失敗: ' + (err && err.message ? err.message : err) + '）';
    Logger.log('[EMSカレンダー] Google同期失敗: ' + (err && err.stack ? err.stack : err));
  }
  ss.toast('EMSカレンダーを更新したで' + calMsg);
}

/** EMSカレンダーの「当日」ハイライトだけを更新（シート再構築なし） */
function refreshEmsCalendarTodayHighlight_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('EMSカレンダー');
  if (!sh || sh.getLastRow() < 3) return;

  _emsLoadTodayMarks_().forEach(m => _emsClearTodayMark_(sh, m));
  _emsClearStaleTodayHighlights_(sh);

  const today = _emsTodayTokyo_();
  const newMarks = [];
  const lastRow = sh.getLastRow();
  let row = 1;

  while (row <= lastRow) {
    const header = _emsParseCalendarMonthHeader_(sh.getRange(row, 1).getDisplayValue());
    if (!header) { row++; continue; }

    row += 2; // 曜日行をスキップ
    while (row <= lastRow) {
      const nextHeader = _emsParseCalendarMonthHeader_(sh.getRange(row, 1).getDisplayValue());
      if (nextHeader) break;
      if (!_emsIsCalendarDayRow_(sh, row)) { row++; continue; }

      const numRow = row;
      for (let c = 1; c <= 7; c++) {
        const v = sh.getRange(numRow, c).getValue();
        if (typeof v !== 'number' || v < 1 || v > 31) continue;
        const d = new Date(header.year, header.monthIndex, v);
        if (d.getMonth() !== header.monthIndex) continue;
        if (!_emsSameDate_(d, today)) continue;
        const rows = _emsFindWeekBlockHeight_(sh, numRow);
        newMarks.push({ row: numRow, col: c, rows: rows });
      }

      row += _emsFindWeekBlockHeight_(sh, numRow);
    }
  }

  newMarks.forEach(m => _emsApplyTodayMark_(sh, m));
  _emsSaveTodayMarks_(newMarks);
}

function Googleカレンダーをシートに同期() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName('Googleカレンダー同期');
  if (!sh) {
    sh = ss.insertSheet('Googleカレンダー同期');
  }

  const calendarName = 'EMS追跡';
  const calendars = CalendarApp.getCalendarsByName(calendarName);

  if (calendars.length === 0) {
    SpreadsheetApp.getUi().alert(`Googleカレンダー「${calendarName}」が見つかりません。`);
    return;
  }

  const cal = calendars[0];

  const startDate = new Date();
  startDate.setMonth(startDate.getMonth() - 1);

  const endDate = new Date();
  endDate.setMonth(endDate.getMonth() + 3);

  const events = cal.getEvents(startDate, endDate);

  const values = [
    ['開始日', '終了日', 'タイトル', '説明', '場所', 'カレンダー名']
  ];

  events.forEach(ev => {
    values.push([
      ev.getStartTime(),
      ev.getEndTime(),
      ev.getTitle(),
      ev.getDescription(),
      ev.getLocation(),
      calendarName
    ]);
  });

  sh.clearContents();

  sh.getRange(1, 1, values.length, values[0].length)
    .setValues(values);

  sh.getRange(1, 1, 1, values[0].length)
    .setBackground('#ffff00')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (values.length > 1) {
    sh.getRange(2, 1, values.length - 1, 2)
      .setNumberFormat('yyyy/m/d');
  }

  sh.autoResizeColumns(1, values[0].length);

  SpreadsheetApp.getActive().toast(
    `Googleカレンダーをシートに同期しました：${events.length}件`
  );
}

function _emsDateOnly_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === 'number' && isFinite(value)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 24 * 60 * 60 * 1000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const raw = String(value || '').normalize('NFKC').trim();
  if (!raw) return null;
  const text = raw
    .replace(/\(.+?\)/g, '')
    .replace(/（.+?）/g, '')
    .replace(/[年月]/g, '/')
    .replace(/日/g, '')
    .replace(/到着/g, '')
    .replace(/着/g, '')
    .replace(/\s+/g, '')
    .trim();

  let m = text.match(/^(\d{2,4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }

  m = text.match(/^(\d{1,2})[\/.-](\d{1,2})$/);
  if (m) {
    const y = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
    return new Date(y, Number(m[1]) - 1, Number(m[2]));
  }

  m = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = text.match(/^(\d{2})(\d{2})(\d{2})$/);
  if (m) return new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = text.match(/(?:^|[^\d])(\d{2,4})[\/.-](\d{1,2})[\/.-](\d{1,2})(?:$|[^\d])/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }

  m = text.match(/(?:^|[^\d])(\d{1,2})[\/.-](\d{1,2})(?:$|[^\d])/);
  if (m) {
    const y = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
    return new Date(y, Number(m[1]) - 1, Number(m[2]));
  }

  return null;
}

function _emsTodayTokyo_() {
  const parts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd').split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function _emsSameDate_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return fmt(a) === fmt(b);
}

function _emsApplyTodayMark_(sh, mark) {
  sh.getRange(mark.row, mark.col)
    .setBackground(EMS_CAL_TODAY.BG)
    .setFontWeight('bold')
    .setFontColor(EMS_CAL_TODAY.FONT);
  sh.getRange(mark.row, mark.col, mark.rows, 1)
    .setBorder(true, true, true, true, null, null, EMS_CAL_TODAY.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function _emsClearTodayMark_(sh, mark) {
  sh.getRange(mark.row, mark.col)
    .setBackground(null)
    .setFontWeight('normal')
    .setFontColor('#70757a');
  sh.getRange(mark.row, mark.col, mark.rows, 1)
    .setBorder(true, true, true, true, null, null, '#d0d7de', SpreadsheetApp.BorderStyle.SOLID);
}

function _emsSaveTodayMarks_(marks) {
  PropertiesService.getDocumentProperties()
    .setProperty(EMS_CAL_TODAY.MARKS_KEY, JSON.stringify(marks || []));
}

function _emsLoadTodayMarks_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(EMS_CAL_TODAY.MARKS_KEY);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function _emsParseCalendarMonthHeader_(text) {
  const m = String(text || '').match(/(\d{4})年(\d{1,2})月/);
  if (!m) return null;
  return { year: Number(m[1]), monthIndex: Number(m[2]) - 1 };
}

function _emsIsCalendarDayRow_(sh, row) {
  const vals = sh.getRange(row, 1, 1, 7).getValues()[0];
  return vals.some(v => typeof v === 'number' && v >= 1 && v <= 31);
}

function _emsFindWeekBlockHeight_(sh, numRow) {
  const maxScan = 30;
  for (let r = numRow + 1; r <= numRow + maxScan; r++) {
    if (sh.getRowHeight(r) === 8) return r - numRow + 1;
  }
  return 3;
}

function _emsClearStaleTodayHighlights_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return;
  const bgs = sh.getRange(1, 1, lastRow, 7).getBackgrounds();
  const target = EMS_CAL_TODAY.BG.toLowerCase();
  for (let r = 0; r < bgs.length; r++) {
    for (let c = 0; c < 7; c++) {
      if (String(bgs[r][c] || '').toLowerCase() !== target) continue;
      const numRow = r + 1;
      const col = c + 1;
      _emsClearTodayMark_(sh, { row: numRow, col: col, rows: _emsFindWeekBlockHeight_(sh, numRow) });
    }
  }
}

function _emsCalendarHasUsefulBg_(color) {
  const c = String(color || '').trim().toLowerCase();
  if (!c) return false;
  const ignored = {
    '#ffffff': true,
    '#fff': true
  };
  return !ignored[c];
}

function _emsCalendarFirstUsefulBg_() {
  for (let i = 0; i < arguments.length; i++) {
    const color = arguments[i];
    if (_emsCalendarHasUsefulBg_(color)) return color;
  }
  return '';
}

function _codePushUnique_(keys, key) {
  if (key && keys.indexOf(key) < 0) keys.push(key);
}

// 商品コードの別名展開
// 例: KRSJCM03-0506_06 / KRSJCM03-0506-06 / KRSJCM03-06 を同一候補にする。
//     KRSJCM03-0506S はセット品なので短縮せず別物のまま。
function codeKeys_(code) {
  const c = normCode_(code);
  if (!c) return [];
  const keys = [c];

  const extended = c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if (extended && (extended[4] === extended[2] || extended[4] === extended[3])) {
    _codePushUnique_(keys, extended[1] + '-' + extended[4]);             // KRSJCM03-06
    _codePushUnique_(keys, extended[1] + '-' + extended[2] + extended[3]); // KRSJCM03-0506
    return keys;
  }

  const pair = c.match(/^(.+)-(\d{2})(\d{2})$/);
  if (pair) {
    _codePushUnique_(keys, pair[1] + '-' + pair[2]); // KRSJCM03-05
    _codePushUnique_(keys, pair[1] + '-' + pair[3]); // KRSJCM03-06
  }

  return keys;
}

// 日付を yyyy-MM-dd に正規化（"26/06/16(火)" みたいな文字列にも対応）
function _emsFmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (typeof v === 'number' && isFinite(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 24 * 60 * 60 * 1000);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  const s = String(v || '').trim().replace(/\(.+?\)/, '');
  const m = s.match(/^(\d{2,4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) { let y = Number(m[1]); if (y < 100) y += 2000;
    return y + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[3]).slice(-2); }
  const jp = s.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (jp) {
    const y = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
    return y + '-' + ('0'+jp[1]).slice(-2) + '-' + ('0'+jp[2]).slice(-2);
  }
  return s;
}

function _emsIsDateKey_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
}

function EMSリスト_安定重複キー_(row) {
  const track = normTrack_(row[12]);                 // M列 EMS番号
  const purchase = (typeof EMS_転送購入Noキー_ === 'function')
    ? EMS_転送購入Noキー_(row[5])                   // F列 購入No
    : String(row[5] || '').trim().replace(/[\s\u3000]+/g, '');
  const code = normCode_(row[8]);                    // I列 商品コード
  const qty = (typeof EMS_転送数量キー_ === 'function')
    ? EMS_転送数量キー_(row[9])                      // J列 数量
    : String(row[9] || '').normalize('NFKC').replace(/,/g, '').replace(/[\s\u3000]+/g, '').trim();
  if (!track || !code || !qty) return '';
  return [track, purchase, code, qty].join('|');
}

function EMSリスト_安定重複表示_(row) {
  return [
    normTrack_(row[12]) || '(EMS番号空)',
    String(row[5] || '').trim() || '(購入No空)',
    normCode_(row[8]) || '(コード空)',
    'x' + String(row[9] || '').trim()
  ].join(' / ');
}

function EMSリスト_重複行を確認して削除() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('EMSリスト');
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('EMSリストが無いで'); return; }

  const vals = sh.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('見出し行(No.)が見つからんで'); return; }

  // EMS番号＋購入No＋商品コード＋数量 でグループ化。
  // 入荷日やコードだけでは、別のEMS箱に入った同一商品まで消えるので使わない。
  const groups = {};
  for (let i = h + 1; i < vals.length; i++) {
    const key = EMSリスト_安定重複キー_(vals[i]);
    if (!key) continue;
    (groups[key] = groups[key] || []).push(i);
  }

  // 統合先の空欄に流し込む列：B C E F G H L M N
  const MERGE_COLS = [2, 3, 5, 6, 7, 8, 12, 13, 14];

  const plans = [];
  Object.keys(groups).forEach(key => {
    const rows = groups[key];
    if (rows.length < 2) return;
    const keep = Math.min(...rows);
    plans.push({ keep: keep, dels: rows.filter(r => r !== keep) });
  });

  if (plans.length === 0) { ui.alert('重複行は見つからんかったで'); return; }

  let preview = '', delCount = 0;
  plans.forEach(p => {
    delCount += p.dels.length;
    preview += `${EMSリスト_安定重複表示_(vals[p.keep])}: 行${p.keep+1}に統合 ← 行${p.dels.map(d=>d+1).join(',')}削除\n`;
  });

  const res = ui.alert(
    '重複行クリーンアップ',
    `重複グループ ${plans.length}件 / 削除 ${delCount}行\n\n` +
    preview.slice(0, 1200) +
    '\n同じEMS番号＋購入No＋商品コード＋数量だけを対象にします。統合先の空欄だけ、削除する行の値で埋めてから消すで。実行する？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { ui.alert('やめといたで'); return; }

  // 統合（keepの空欄だけ埋める）
  plans.forEach(p => {
    p.dels.forEach(d => {
      MERGE_COLS.forEach(col => {
        if (isBlank_(vals[p.keep][col-1]) && !isBlank_(vals[d][col-1])) {
          sh.getRange(p.keep + 1, col).setValue(vals[d][col-1]);
          vals[p.keep][col-1] = vals[d][col-1];
        }
      });
    });
  });

  // 行削除（下から）
  const delRows = [];
  plans.forEach(p => p.dels.forEach(d => delRows.push(d)));
  [...new Set(delRows)].sort((a,b)=>b-a).forEach(r => sh.deleteRow(r + 1));

  ui.alert(`完了: ${delRows.length}行 削除したで`);
}

function EMSリスト_下に混ざった古い行を確認して削除() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('EMSリスト');
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('EMSリストが無いで'); return; }

  const vals = sh.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('見出し行(No.)が見つからんで'); return; }

  let latestSeen = '';
  const targets = [];
  const firstByStableKey = {};
  for (let i = h + 1; i < vals.length; i++) {
    const key = EMSリスト_安定重複キー_(vals[i]);
    if (key && firstByStableKey[key] === undefined) firstByStableKey[key] = i;
  }

  for (let i = h + 1; i < vals.length; i++) {
    const arr = _emsFmtDate_(vals[i][1]); // B列 入荷日
    if (!_emsIsDateKey_(arr)) continue;
    if (latestSeen && arr < latestSeen) {
      const key = EMSリスト_安定重複キー_(vals[i]);
      const first = key ? firstByStableKey[key] : undefined;
      if (first !== undefined && first < i) {
        targets.push({
          row: i + 1,
          firstRow: first + 1,
          arr: arr,
          purchase: String(vals[i][5] || ''),
          code: normCode_(vals[i][8]),
          qty: String(vals[i][9] || ''),
          track: normTrack_(vals[i][12])
        });
      }
      continue;
    }
    if (arr > latestSeen) latestSeen = arr;
  }

  if (targets.length === 0) {
    ui.alert('下に混ざった古い行は見つからんかったで');
    return;
  }

  const preview = targets
    .slice(0, 40)
    .map(t => `行${t.row}: ${t.arr} ${t.purchase} ${t.code} x${t.qty} ${t.track}（上の行${t.firstRow}と同一）`)
    .join('\n');
  const more = targets.length > 40 ? `\n...ほか ${targets.length - 40}行` : '';
  const res = ui.alert(
    '下に混ざった古いEMS行を削除',
    `新しい入荷日の下にある古い行のうち、同じEMS番号＋購入No＋商品コード＋数量が上にもある行だけ ${targets.length}行 見つけました。\n\n${preview}${more}\n\n削除する？`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { ui.alert('やめといたで'); return; }

  targets.map(t => t.row).sort((a,b)=>b-a).forEach(row => sh.deleteRow(row));
  ui.alert(`完了: ${targets.length}行 削除したで`);
}
// updated menu setup
