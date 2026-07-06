// ====== EMSリスト：P列(注文番号)を受注明細との照合で自動記入 ======
// 引き当てファイル(GoQ受注明細)の「取り寄せ」行と、EMSリストの各行(商品コード＋発注日)を
// 発注日ベースのFIFOで突き合わせて、P列に受注番号を自動で書き込む。
//
// 書式: 全量1注文 → 「10117052」 / 分割 → 「10117060:3, 10117052:1」
// ルール:
//   ・P列に既に値がある行は触らない(手動記入・修正が常に優先)
//   ・「注文日時 ≦ 発注日」の取り寄せ注文だけを候補にする(発注より後の注文は在庫買いに紐付けない)
//   ・数量の残り(どの注文にも割り当たらない分)は書かない＝在庫扱い
//   ・分割/一部在庫の行はP列セルを薄黄にして目視確認しやすくする

const EMS_JUCHU_CFG = {
  引当ファイルID: '15n4snWF2lPOggm4rlmp-cux_3qlV7VpSh3L2o_41qHo', // GoQ受注→EMS引き当てファイル
  受注シート: '受注明細',
  EMSシート: 'EMSリスト',
  EMS_ヘッダー行: 6,   // 見出し(No./商品コード/照合キー…)がある行
  EMS_データ開始行: 7,
  色_要確認: '#fff2cc'  // 分割・一部在庫のP列セル
};

// メニューは【統合】onOpen管理.jsのEMS自社用に追加済み → 'EMSリスト_注文番号自動記入'

function EMSリスト_注文番号自動記入() {
  const cfg = EMS_JUCHU_CFG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ems = ss.getSheetByName(cfg.EMSシート);
  if (!ems) { SpreadsheetApp.getUi().alert('EMSリストシートが見つかりません。'); return; }

  // ---- 受注明細(別ファイル)から取り寄せ行を読む ----
  let recv;
  try {
    recv = SpreadsheetApp.openById(cfg.引当ファイルID).getSheetByName(cfg.受注シート);
  } catch (e) {
    SpreadsheetApp.getUi().alert('引き当てファイルが開けません:\n' + e.message);
    return;
  }
  if (!recv) { SpreadsheetApp.getUi().alert('引き当てファイルに「' + cfg.受注シート + '」がありません。'); return; }

  const lines = EMSJ_受注取り寄せ行_(recv); // [{ban, keys:Set, need, date}]
  if (!lines.length) {
    SpreadsheetApp.getUi().alert('受注明細に取り寄せの注文が見つかりません。\n先に引き当てファイルでCSV取込をしてください。');
    return;
  }
  const 受注最古 = lines.reduce((m, l) => Math.min(m, l.date.getTime()), Infinity); // ※現役の注文だけで計算

  // 消込台帳の「出荷済み」行も候補に加える(発送済みでも「この箱の商品は誰の分だったか」をP列に残すため)
  // shipped=true の行は日付フィルタを通さず、割り当て順は現役の注文より先(物理的に先に取っていった分)
  const 台帳lines = EMSJ_台帳出荷済み行_(recv.getParent());
  台帳lines.forEach(l => { l.seq = lines.length + l.seq; lines.push(l); });

  // key(正規化コード) → その商品を待っている注文行(出荷済みが先、あとは古い順)
  const byKey = {};
  lines.forEach(l => l.keys.forEach(k => (byKey[k] = byKey[k] || []).push(l)));
  Object.keys(byKey).forEach(k => byKey[k].sort((a, b) =>
    ((a.shipped ? 0 : 1) - (b.shipped ? 0 : 1)) || a.date - b.date || a.seq - b.seq));

  // ---- EMSリストを読む ----
  const lastRow = ems.getLastRow();
  if (lastRow < cfg.EMS_データ開始行) { SpreadsheetApp.getActive().toast('EMSリストにデータがありません。'); return; }

  const head = ems.getRange(cfg.EMS_ヘッダー行, 1, 1, ems.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim());
  const colP = head.indexOf('注文番号') + 1;           // 1始まり
  const colPurchase = head.indexOf('購入No.') >= 0 ? head.indexOf('購入No.') + 1 : 6; // F列
  const colCode = head.indexOf('商品コード') >= 0 ? head.indexOf('商品コード') + 1 : 9; // I列
  const colQty = head.indexOf('数量') >= 0 ? head.indexOf('数量') + 1 : 10;             // J列
  if (colP === 0) { SpreadsheetApp.getUi().alert('EMSリストの' + cfg.EMS_ヘッダー行 + '行目に「注文番号」見出しがありません。P列に見出しを入れてください。'); return; }

  const n = lastRow - cfg.EMS_データ開始行 + 1;
  const purchase = ems.getRange(cfg.EMS_データ開始行, colPurchase, n, 1).getDisplayValues();
  const codes = ems.getRange(cfg.EMS_データ開始行, colCode, n, 1).getDisplayValues();
  const qtys = ems.getRange(cfg.EMS_データ開始行, colQty, n, 1).getValues();
  const pVals = ems.getRange(cfg.EMS_データ開始行, colP, n, 1).getDisplayValues();

  const rows = [];
  for (let i = 0; i < n; i++) {
    const pno = String(purchase[i][0] || '').trim();
    const code = String(codes[i][0] || '').trim();
    if (!pno || !code) continue;
    const m = pno.match(/^(\d{4})(\d{2})(\d{2})/); // 購入No.先頭8桁＝発注日
    if (!m) continue;
    const 発注日末 = new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59); // その日の注文までOK
    rows.push({
      i, 発注日末,
      keys: codeKeys_(code),                 // Project_19共通の別名展開
      qty: Number(qtys[i][0]) || 0,
      p: String(pVals[i][0] || '').trim()
    });
  }

  // ---- 1周目: 既にP列に入っている分を注文の必要数から差し引く(二重割当防止) ----
  rows.forEach(r => {
    if (!r.p) return;
    EMSJ_P列パース_(r.p, r.qty).forEach(e => {
      let left = e.qty;
      for (const k of r.keys) {
        for (const l of (byKey[k] || [])) {
          if (left <= 0) break;
          if (l.ban !== e.ban || l.need <= 0) continue;
          const take = Math.min(left, l.need);
          l.need -= take; left -= take;
        }
        if (left <= 0) break;
      }
      // 対応する注文行が受注明細に無い(古い注文等)場合はそのまま(問題なし)
    });
  });

  // ---- 2周目: 空欄の行へ発注日順にFIFOで割り当て ----
  const target = rows.filter(r => !r.p && r.qty > 0 && r.発注日末.getTime() >= 受注最古)
    .sort((a, b) => a.発注日末 - b.発注日末 || a.i - b.i);

  let 記入 = 0, 分割 = 0, 全量在庫 = 0;
  const writes = [];   // {i, text, warn}
  target.forEach(r => {
    // この行のコードを待つ注文(出荷済みが先→古い順・重複行を除いてマージ)
    const seen = new Set(), cand = [];
    r.keys.forEach(k => (byKey[k] || []).forEach(l => { if (!seen.has(l.seq)) { seen.add(l.seq); cand.push(l); } }));
    cand.sort((a, b) => ((a.shipped ? 0 : 1) - (b.shipped ? 0 : 1)) || a.date - b.date || a.seq - b.seq);

    let left = r.qty;
    const got = []; // {ban, take}
    for (const l of cand) {
      if (left <= 0) break;
      if (l.need <= 0) continue;
      if (!l.shipped && l.date.getTime() > r.発注日末.getTime()) continue; // 現役の注文だけ「注文日≦発注日」を課す
      const take = Math.min(left, l.need);
      l.need -= take; left -= take;
      const prev = got.find(g => g.ban === l.ban);
      if (prev) prev.take += take; else got.push({ ban: l.ban, take });
    }
    if (!got.length) { 全量在庫++; return; } // 誰の分でもない=在庫買い。空欄のまま

    const full = (got.length === 1 && left === 0); // 全量1注文
    const text = full ? got[0].ban : got.map(g => g.ban + ':' + g.take).join(', ');
    writes.push({ i: r.i, text, warn: !full });
    記入++; if (!full) 分割++;
  });

  // ---- 書き込み(値は列ごと一括・色は変更セルだけ) ----
  if (writes.length) {
    const col = pVals.map(v => [v[0]]); // 既存値を保持したまま列を書き戻す
    writes.forEach(w => { col[w.i][0] = w.text; });
    ems.getRange(cfg.EMS_データ開始行, colP, n, 1).setValues(col);
    const warnCells = writes.filter(w => w.warn)
      .map(w => ems.getRange(cfg.EMS_データ開始行 + w.i, colP).getA1Notation());
    if (warnCells.length) ems.getRangeList(warnCells).setBackground(cfg.色_要確認);
  }

  SpreadsheetApp.getActive().toast(
    '注文番号の自動記入: ' + 記入 + '行' +
    (分割 ? '（うち分割/一部在庫 ' + 分割 + '行=薄黄）' : '') +
    ' / 該当注文なし(在庫扱い) ' + 全量在庫 + '行' +
    ' / 既存スキップ ' + rows.filter(r => r.p).length + '行',
    'P列自動記入', 8);
}

// 受注明細から「取り寄せ」行を読む → [{ban, keys, need, date, seq}]
function EMSJ_受注取り寄せ行_(sh) {
  const vals = sh.getDataRange().getValues();
  let hr = -1;
  for (let i = 0; i < Math.min(vals.length, 50); i++) {
    if (vals[i].some(c => String(c).trim() === '受注番号')) { hr = i; break; }
  }
  if (hr < 0) return [];
  const head = vals[hr].map(v => String(v || '').trim());
  const f = (...names) => { for (const nm of names) { const i = head.indexOf(nm); if (i >= 0) return i; } return -1; };
  const c番号 = f('受注番号'), cコード = f('商品コード'), cSKU = f('商品SKU', 'SKU'),
        c個数 = f('個数'), c選択肢 = f('項目・選択肢', '項目選択肢'), c日時 = f('注文日時');
  if (c番号 < 0 || cコード < 0) return [];

  const out = [];
  for (let i = hr + 1; i < vals.length; i++) {
    const r = vals[i];
    const ban = String(r[c番号] || '').trim();
    if (!ban) continue;
    const opt = String(c選択肢 >= 0 ? r[c選択肢] || '' : '');
    if (!(opt.indexOf('取り寄せ') >= 0 || opt.indexOf('取寄') >= 0)) continue; // 取り寄せのみ対象
    const qty = Number(c個数 >= 0 ? r[c個数] : 0) || 0;
    if (qty <= 0) continue;
    const date = EMSJ_日時_(c日時 >= 0 ? r[c日時] : '');
    if (!date) continue;

    // 照合キー候補: SKU末尾の英字を落としたもの / SKU / 商品コード → 別名展開
    const sku = String(cSKU >= 0 ? r[cSKU] || '' : '').trim();
    const code = String(r[cコード] || '').trim();
    const keys = new Set();
    [sku.replace(/[A-Za-z]$/, ''), sku, code].forEach(v => {
      if (v) codeKeys_(v).forEach(k => keys.add(k));
    });
    if (!keys.size) continue;
    out.push({ ban, keys, need: qty, date, seq: out.length });
  }
  return out;
}

// 引き当てファイルの消込台帳から「出荷済み」行を読む → [{ban, keys, need, date, seq, shipped:true}]
// 発送済みで受注明細から消えた注文。60日以内のものだけ(古い出荷が古い箱以外に付かないように)
function EMSJ_台帳出荷済み行_(ss) {
  const sh = ss.getSheetByName('消込台帳');
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(); // A受注番号 B商品コード C SKU D個数 E入荷日 F状態 G初回 H消滅日
  const limit = new Date(); limit.setDate(limit.getDate() - 60);
  const out = [];
  vals.forEach(r => {
    if (String(r[5] || '').trim().indexOf('出荷済み') !== 0) return;
    const ban = String(r[0] || '').trim();
    const qty = Number(r[3]) || 0;
    if (!ban || qty <= 0) return;
    const base = r[4] || r[7]; // 入荷日→なければ消滅日
    const d = base instanceof Date ? base : new Date(String(base || ''));
    if (!isNaN(d.getTime()) && d.getTime() < limit.getTime()) return;
    const code = String(r[1] || '').trim(), sku = String(r[2] || '').trim();
    const keys = new Set();
    [sku.replace(/[A-Za-z]$/, ''), sku, code].forEach(v => { if (v) codeKeys_(v).forEach(k => keys.add(k)); });
    if (!keys.size) return;
    out.push({ ban, keys, need: qty, date: isNaN(d.getTime()) ? new Date(0) : d, seq: out.length, shipped: true });
  });
  return out;
}

// P列の値をパース: 「10117052」(全量) / 「10117060:3, 10117052:1」 → [{ban, qty}]
function EMSJ_P列パース_(text, rowQty) {
  const out = [];
  String(text).split(/[,、]/).forEach(part => {
    const m = String(part).trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
    if (!m) return;
    out.push({ ban: m[1], qty: m[2] ? Number(m[2]) : rowQty });
  });
  return out;
}

// 注文日時をDateへ(Date/文字列どちらでも)
function EMSJ_日時_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s.replace(/\//g, '-'));
  return isNaN(d.getTime()) ? null : d;
}
