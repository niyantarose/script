// ====== 手動運用メニュー ======
function Excel同期メニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('手動運用')
    .addItem('EMSカレンダーシートを更新', 'buildEmsCalendarSheet')
    .addItem('GoogleカレンダーへEMSを反映', 'syncEmsCalendar')
    .addItem('EMSリスト：重複行を確認して削除', 'EMSリスト_重複行を確認して削除')
    .addItem('EMSリスト：下に混ざった古い行を確認して削除', 'EMSリスト_下に混ざった古い行を確認して削除')
    .addSeparator()
    .addItem('商品マスタ：足りないデータだけ発注から補完', '商品マスタ_足りないデータだけ発注から補完')
    .addItem('商品マスタ：既存データを発注から更新（重さ等）', '商品マスタ_既存データを発注から補完更新')
    .addItem('商品マスタ：重複行を確認して削除', '商品マスタ_重複行を確認して削除')
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
  return String(v || '').trim().toUpperCase().replace(/_/g, '-');
}
function normTrack_(v) {
  return String(v || '').replace(/[\s\u3000]/g, '').toUpperCase();
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '' || String(v).trim() === '　';
}

// ====== EMSリスト → Googleカレンダー同期 ======
const EMS_CAL_CFG = {
  CALENDAR_NAME: 'EMS追跡',
  MARKER: '[EMS-SYNC]',  // スクリプトが作ったイベントの目印
  COLORS: {
    '未着':       CalendarApp.EventColor.RED,
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
    const ship = r[2], arrive = r[4];             // C列 発送日 / E列 到着日
    if (!(ship instanceof Date)) continue;        // 発送日が無い行はスキップ
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
  const today = new Date(); today.setHours(0, 0, 0, 0);   // ★追加
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
    const color = (endBase < today)                        // ★到着済みはグレー
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
  let header = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') { header = i; break; }
  }
  if (header < 0) return {};

  // M列（EMS番号）の背景色＝EMSリストの箱グループ色を帯にも使う
  const emsBgs = sh.getRange(1, 13, vals.length, 1).getBackgrounds();

  const groups = {};
  for (let i = header + 1; i < vals.length; i++) {
    const r = vals[i];
    const track = normTrack_(r[12]);
    if (!track) continue;
    const ship = r[2], arrive = r[4];
    if (!(ship instanceof Date)) continue;
    const key = track + '|' + String(r[13]);
    if (!groups[key]) {
      groups[key] = {
        track: track, box: r[13], ship: ship,
        arrive: (arrive instanceof Date) ? arrive : null,
        status: String(r[6] || '').trim(),
        color: emsBgs[i][0],
        items: [], qty: 0
      };
    }
    groups[key].items.push(`${r[8]} ×${r[9]}（${r[10]}）`);
    groups[key].qty += Number(r[9]) || 0;
    if (String(r[6]).trim() === '未着') groups[key].status = '未着';
  }
  return groups;
}

// 表示する期間（今日を基準に前後何ヶ月か）
const EMS_CAL_VIEW = { MONTHS_BACK: 1, MONTHS_FWD: 1 };

function buildEmsCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groups = getEmsBoxGroups_(ss);
  let boxes = Object.keys(groups).map(k => groups[k]);
  if (boxes.length === 0) { ss.toast('EMSデータがありません'); return; }

  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const STYLE = {
    '未着':         { bg: '#d93025', font: '#ffffff' },
    '輸送中':       { bg: '#f29900', font: '#ffffff' },
    '在庫反映済み': { bg: '#b7dec7', font: '#3c4043' },
    'default':      { bg: '#dadce0', font: '#3c4043' }
  };
  const PAST_STYLE = { bg: '#e8eaed', font: '#9aa0a6' }; // 到着済みのグレーアウト

  const today = new Date(); today.setHours(0, 0, 0, 0);

  // 表示ウィンドウ（先月初〜2ヶ月先の月末）
  const winStart = new Date(today.getFullYear(), today.getMonth() - EMS_CAL_VIEW.MONTHS_BACK, 1);
  const winEnd = new Date(today.getFullYear(), today.getMonth() + EMS_CAL_VIEW.MONTHS_FWD + 1, 0);

  // 箱の整形＋ウィンドウ外を除外
  boxes.forEach(b => { b.end = b.arrive || b.ship; });
  boxes = boxes.filter(b => b.end >= winStart && b.ship <= winEnd);

  boxes.forEach(b => {
    b.isPast = b.end < today;
    if (b.isPast) {
      // 到着済みは今まで通りグレー
      b.style = PAST_STYLE;
    } else {
      // EMSリストのM列の色をそのまま帯に使う（色が無ければステータス色にフォールバック）
      const hasColor = b.color && String(b.color).toLowerCase() !== '#ffffff';
      b.style = hasColor
        ? { bg: b.color, font: '#3c4043' }
        : (STYLE[b.status] || STYLE['default']);
    }
    b.note = `【${b.status || '?'}】${b.track} Box${b.box}\n` +
             `発送: ${fmt(b.ship)} → 到着: ${fmt(b.end)}\n` +
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
    if (days >= 4) return `【${b.status || '?'}】${b.track} B${b.box}（${b.qty}点）`;
    if (days >= 2) return `【${b.status || '?'}】${b.track} B${b.box}`;
    return `【${b.status || '?'}】${b.track}`;
  };

  let sh = ss.getSheetByName('EMSカレンダー');
  if (!sh) sh = ss.insertSheet('EMSカレンダー');
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearNotes();

  const youbi = ['日', '月', '火', '水', '木', '金', '土'];
  let row = 1;
  let cur = new Date(winStart);

  while (cur <= winEnd) {
    const y = cur.getFullYear(), mo = cur.getMonth();
    const monthStart = new Date(y, mo, 1);
    const daysInMonth = new Date(y, mo + 1, 0).getDate();
    const monthEnd = new Date(y, mo, daysInMonth);
    const blockStart = row;

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
      for (let c = 0; c < 7; c++) {
        const d = new Date(weekStart); d.setDate(d.getDate() + c);
        const cell = sh.getRange(numRow, c + 1);
        if (d.getMonth() === mo) {
          cell.setValue(d.getDate());
          if (fmt(d) === fmt(today)) {
            cell.setBackground('#fde293').setFontWeight('bold').setFontColor('#3c4043');
          }
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

      segs.forEach(x => {
        const c1 = Math.floor((x.s - weekStart) / 86400000) + 1;
        const c2 = Math.floor((x.e - weekStart) / 86400000) + 1;
        const days = c2 - c1 + 1;
        const range = sh.getRange(row + x.b.lane, c1, 1, days);
        if (days > 1) range.merge();
        range.setValue(bandLabel(x.b, days))
          .setBackground(x.b.style.bg).setFontColor(x.b.style.font)
          .setFontSize(12)   // ★12ptに
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

    row++;
    cur = new Date(y, mo + 1, 1);
  }

 for (let c = 1; c <= 7; c++) sh.setColumnWidth(c, 180);  // ★180に
  sh.setHiddenGridlines(true);
  ss.toast('EMSカレンダーを更新したで');
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

// 商品コードの別名展開（KRSJCM03-0506-06 → KRSJCM03-06 も同一とみなす）
function codeKeys_(code) {
  const c = normCode_(code);
  const keys = [c];
  const m = c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if (m && (m[4] === m[2] || m[4] === m[3])) {
    keys.push(m[1] + '-' + m[4]);   // 短縮形を別名として追加
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

function EMSリスト_重複行を確認して削除() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('EMSリスト');
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('EMSリストが無いで'); return; }

  const vals = sh.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('見出し行(No.)が見つからんで'); return; }

  // 入荷日＋商品コード＋数量 でグループ化
  const groups = {};
  for (let i = h + 1; i < vals.length; i++) {
    const code = normCode_(vals[i][8]);     // I列 商品コード
    if (!code) continue;
    const arr = _emsFmtDate_(vals[i][1]);   // B列 入荷日
    const qty = String(vals[i][9]);         // J列 数量
    const key = arr + '|' + code + '|' + qty;
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
    preview += `${normCode_(vals[p.keep][8])}: 行${p.keep+1}に統合 ← 行${p.dels.map(d=>d+1).join(',')}削除\n`;
  });

  const res = ui.alert(
    '重複行クリーンアップ',
    `重複グループ ${plans.length}件 / 削除 ${delCount}行\n\n` +
    preview.slice(0, 1200) +
    '\n統合先の空欄だけ、削除する行の値で埋めてから消すで。実行する？',
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
  for (let i = h + 1; i < vals.length; i++) {
    const arr = _emsFmtDate_(vals[i][1]); // B列 入荷日
    if (!_emsIsDateKey_(arr)) continue;
    if (latestSeen && arr < latestSeen) {
      targets.push({
        row: i + 1,
        arr: arr,
        purchase: String(vals[i][5] || ''),
        code: normCode_(vals[i][8]),
        qty: String(vals[i][9] || ''),
        track: normTrack_(vals[i][12])
      });
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
    .map(t => `行${t.row}: ${t.arr} ${t.purchase} ${t.code} x${t.qty} ${t.track}`)
    .join('\n');
  const more = targets.length > 40 ? `\n...ほか ${targets.length - 40}行` : '';
  const res = ui.alert(
    '下に混ざった古いEMS行を削除',
    `新しい入荷日の下にある古い行を ${targets.length}行 見つけました。\n\n${preview}${more}\n\n削除する？`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) { ui.alert('やめといたで'); return; }

  targets.map(t => t.row).sort((a,b)=>b-a).forEach(row => sh.deleteRow(row));
  ui.alert(`完了: ${targets.length}行 削除したで`);
}
