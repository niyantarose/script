// ===== P列自動記入(引当ファイル側から発注共有ファイルのEMSリストへ書き込む) =====
// 発注共有ファイル側のスクリプトと同じロジック。受注明細(現役)＋消込台帳(出荷済み)を
// EMSリストの各行(商品コード＋発注日=購入No.先頭8桁)と照合して、P列に受注番号を書く。
// ②引き当て実行の最初にも自動で走るので、通常は手で押す必要はない。
//
// ルール:
//   ・P列に既に値がある行は触らない(手動修正が常に優先)
//   ・出荷済み(消込台帳)を最優先で割り当て(物理的に先に取っていった分)
//   ・現役の注文は「注文日≦発注日」のものだけ確定(後から来た注文は引き当て時のFIFOに任せる)
//   ・全量1注文=番号だけ / 分割=「番号:個数」カンマ区切り(セルは薄黄で要確認マーク)

function P列に注文番号を自動記入(){
  const r=発注共有P列記入_();
  if(r.error){ SpreadsheetApp.getUi().alert(r.error); return; }
  SpreadsheetApp.getActive().toast(
    'P列自動記入: '+r.記入+'行'+(r.分割?'（分割/一部在庫 '+r.分割+'行=薄黄）':'')+
    ' / 該当注文なし(在庫扱い) '+r.在庫+'行 / 既存スキップ '+r.既存+'行','📝P列',8);
}

function 発注共有P列記入_(){
  const cfg=P_KAKUTEI_CFG, ss=SpreadsheetApp.getActive();
  let ems;
  try{ ems=SpreadsheetApp.openById(cfg.発注共有ID).getSheetByName(cfg.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!ems) return {error:'発注共有ファイルに「'+cfg.シート+'」がありません'};
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return {error:'「'+HIKIATE_CFG.受注+'」タブがありません'};

  // ---- 現役の取り寄せ行(受注明細) ----
  const M=列マップ_(recv), R=recv.getDataRange().getValues();
  const head=recv.getRange(M.hr,1,1,recv.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const c日時=head.indexOf('注文日時');
  const 日時_=v=>{ if(v instanceof Date) return isNaN(v.getTime())?null:v;
    const s=String(v||'').trim(); if(!s) return null;
    const d=new Date(s.replace(/\//g,'-')); return isNaN(d.getTime())?null:d; };
  const キー展開_=(sku,code)=>{ const keys=new Set();
    受注候補コード_(sku,code).forEach(v=>{ if(v) codeKeys_(v).forEach(k=>keys.add(k)); });
    return keys; };
  const lines=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i]; const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const qty=Number(row[M.個数])||0; if(qty<=0) continue;
    const date=c日時>=0? 日時_(row[c日時]) : null; if(!date) continue;
    const keys=キー展開_(M.SKU>=0?String(row[M.SKU]||'').trim():'', String(row[M.コード]||'').trim());
    if(!keys.size) continue;
    lines.push({ban, keys, need:qty, date, seq:lines.length, shipped:false});
  }
  if(!lines.length) return {error:'受注明細に取り寄せの注文がありません。先にCSV取込をしてください。'};
  const 受注最古=lines.reduce((m,l)=>Math.min(m,l.date.getTime()),Infinity); // ※現役の注文だけで計算

  // ---- 出荷済み(消込台帳)も候補に(発送済みでも「誰の分だったか」をP列に残す) ----
  消込台帳_出荷済み行_().forEach(t=>{
    const keys=キー展開_(t.sku, t.code); if(!keys.size) return;
    lines.push({ban:t.ban, keys, need:t.qty, date:new Date(0), seq:lines.length, shipped:true});
  });
  // 引当履歴の「在庫反映済み/過去取込」分も必要数から差し引き、古い箱で割当済みの注文を新しい箱に二重記入しない。
  try{ 引当履歴_需要を差し引く_(lines); }catch(e){}

  const byKey={};
  lines.forEach(l=>l.keys.forEach(k=>(byKey[k]=byKey[k]||[]).push(l)));
  const 順_=(a,b)=>((a.shipped?0:1)-(b.shipped?0:1)) || a.date-b.date || a.seq-b.seq; // 出荷済み→古い注文
  Object.keys(byKey).forEach(k=>byKey[k].sort(順_));

  // ---- EMSリストを読む(見出し名で列を特定) ----
  const hr=cfg.ヘッダー行, last=ems.getLastRow();
  if(last<=hr) return {error:'EMSリストにデータがありません'};
  const eh=ems.getRange(hr,1,1,ems.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const colP=eh.indexOf('注文番号')+1;
  const colPu=(eh.indexOf('購入No.')>=0? eh.indexOf('購入No.'):5)+1;
  const colC=(eh.indexOf('商品コード')>=0? eh.indexOf('商品コード'):8)+1;
  const colQ=(eh.indexOf('数量')>=0? eh.indexOf('数量'):9)+1;
  if(colP===0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};

  const n=last-hr;
  const pu=ems.getRange(hr+1,colPu,n,1).getDisplayValues();
  const cd=ems.getRange(hr+1,colC,n,1).getDisplayValues();
  const qy=ems.getRange(hr+1,colQ,n,1).getValues();
  const pv=ems.getRange(hr+1,colP,n,1).getDisplayValues();
  const rows=[];
  for(let i=0;i<n;i++){
    const pno=String(pu[i][0]||'').trim(), code=String(cd[i][0]||'').trim();
    if(!pno||!code) continue;
    const m=pno.match(/^(\d{4})(\d{2})(\d{2})/); if(!m) continue; // 購入No.先頭8桁=発注日
    rows.push({ i, 末:new Date(+m[1],+m[2]-1,+m[3],23,59,59),
      keys:codeKeys_(code), 号:月号_(normCode_(code)), qty:Number(qy[i][0])||0, p:String(pv[i][0]||'').trim() });
  }

  // ---- 1周目: 既にP列にある分を必要数から差し引く(二重割当防止) ----
  const parse_=(t,rq)=>{ const out=[];
    String(t).split(/[,、]/).forEach(p=>{ const m=String(p).trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
      if(m) out.push({ban:m[1], qty:m[2]?Number(m[2]):rq}); });
    return out; };
  rows.forEach(r=>{
    if(!r.p) return;
    parse_(r.p,r.qty).forEach(e=>{ let left=e.qty;
      for(const k of r.keys){
        for(const l of (byKey[k]||[])){ if(left<=0) break;
          if(l.ban!==e.ban||l.need<=0) continue;
          const t=Math.min(left,l.need); l.need-=t; left-=t; }
        if(left<=0) break; } });
  });

  // ---- 2周目: 空欄の行へ発注日順にFIFOで記入 ----
  const target=rows.filter(r=>!r.p && r.qty>0 && r.末.getTime()>=受注最古).sort((a,b)=>a.末-b.末||a.i-b.i);
  let 記入=0, 分割=0, 在庫=0; const writes=[];
  target.forEach(r=>{
    const seen=new Set(), cand=[];
    r.keys.forEach(k=>(byKey[k]||[]).forEach(l=>{ if(!seen.has(l.seq)){ seen.add(l.seq); cand.push(l); } }));
    cand.sort(順_);
    let left=r.qty; const got=[];
    for(const l of cand){
      if(left<=0) break;
      if(l.need<=0) continue;
      if(!l.shipped && l.date.getTime()>r.末.getTime()) continue; // 現役の注文だけ「注文日≦発注日」
      if(!l.shipped && r.号 && 期待号_(l.date)!==r.号) continue;   // 月号付き(定期購読)は「注文月+1=その号」の注文だけ
      const t=Math.min(left,l.need); l.need-=t; left-=t;
      const prev=got.find(g=>g.ban===l.ban);
      if(prev) prev.take+=t; else got.push({ban:l.ban, take:t});
    }
    if(!got.length){ 在庫++; return; }
    const full=(got.length===1 && left===0);
    writes.push({ i:r.i, text: full? got[0].ban : got.map(g=>g.ban+':'+g.take).join(', '), warn:!full });
    記入++; if(!full) 分割++;
  });

  if(writes.length){
    const col=pv.map(v=>[v[0]]); // 既存値を保ったまま列ごと書き戻す
    writes.forEach(w=>{ col[w.i][0]=w.text; });
    ems.getRange(hr+1,colP,n,1).setValues(col);
    const warn=writes.filter(w=>w.warn).map(w=>ems.getRange(hr+1+w.i,colP).getA1Notation());
    if(warn.length) ems.getRangeList(warn).setBackground('#fff2cc');
  }
  return {記入, 分割, 在庫, 既存:rows.filter(r=>r.p).length};
}
