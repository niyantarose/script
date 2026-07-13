// ===== P列確定引き当て: 発注共有ファイルのEMSリストP列(注文番号)を読む =====
// 発注時点で「この商品は誰の注文の分か」がP列に記入されている(自動記入＋手動修正)。
// 引き当て実行はこの名指しを最優先で割り当て、名指しのない分だけ従来のコードFIFOで配分する。
// → 商品コードの表記ゆれや同一コード複数注文でも、紐付けが「推測」でなく「確定」になる。

const P_KAKUTEI_CFG = {
  発注共有ID: '1yDO8ae4TNeJceSpPhuALRF48PKX55PJ_nJ5ACwAokes', // 発注共有ファイル
  シート: 'EMSリスト',
  ヘッダー行: 6 // No./ステータス列/商品コード/数量/照合キー/注文番号 の見出し行
};

// 発注共有ファイルは1回の実行で何度も開くと遅い(②はP列記入・確定・履歴で複数回開いていた)。
// 実行中は同じハンドルを使い回してopenByIdの回数を減らす(GASのグローバルは実行ごとにリセットされる)。
var _発注共有SS_キャッシュ=null;
function 発注共有を開く_(){
  if(!_発注共有SS_キャッシュ) _発注共有SS_キャッシュ=SpreadsheetApp.openById(P_KAKUTEI_CFG.発注共有ID);
  return _発注共有SS_キャッシュ;
}

// 到着済のEMSリスト行のP列を読んで、
// 受注番号→[{key:正規化コード, qty:名指し個数, arrival:EMS到着日, ems:EMS番号}] を返す
// P列の書式: 「10117052」(行の全量) / 「10117060:3, 10117052:1」(分割)
function P列確定マップ_(){
  const cfg=P_KAKUTEI_CFG;
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(e){ return {}; } // 開けない(権限・ID変更等)時は確定なし=従来FIFOのみで動く
  if(!sh) return {};
  const last=sh.getLastRow(); if(last<=cfg.ヘッダー行) return {};
  const head=sh.getRange(cfg.ヘッダー行,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const f=n=>head.indexOf(n);
  const cSt=f('ステータス列'), cCode=f('商品コード'), cQty=f('数量'), cP=f('注文番号'),
        cArrival=f('EMS到着日'), cEms=f('EMS番号');
  if(cP<0 || cCode<0) return {};
  const vals=sh.getRange(cfg.ヘッダー行+1,1,last-cfg.ヘッダー行,sh.getLastColumn()).getDisplayValues();
  const map={};
  vals.forEach(r=>{
    if(cSt>=0 && String(r[cSt]||'').trim()!=='到着済') return; // 実際に届いている分だけが引き当て対象
    if(cEms<0 || !実EMS番号_(r[cEms])) return; // 棚卸箱・EMS番号空欄は実物の供給ではない
    const p=String(r[cP]||'').trim(); if(!p) return;
    const key=normCode_(r[cCode]); if(!key) return;
    const rowQty= cQty>=0? (Number(String(r[cQty]||'').replace(/[^\d.]/g,''))||0) : 0;
    p.split(/[,、]/).forEach(part=>{
      const m=String(part).trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
      if(!m) return;
      (map[m[1]]=map[m[1]]||[]).push({
        key,
        qty: m[2]? Number(m[2]) : rowQty,
        arrival: cArrival>=0? String(r[cArrival]||'').trim() : '',
        ems: cEms>=0? String(r[cEms]||'').trim() : ''
      });
    });
  });
  return map;
}
