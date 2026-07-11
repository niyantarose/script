// ===== P列自動記入(引当ファイル側から発注共有ファイルのEMSリストへ書き込む) =====
// 発注共有ファイル側のスクリプトと同じロジック。受注明細(現役)＋消込台帳(出荷済み)を
// EMSリストの各行(商品コード＋発注日=購入No.先頭8桁)と照合して、P列に受注番号を書く。
// ②引き当て実行の最初にも自動で走るので、通常は手で押す必要はない。
//
// ルール:
//   ・P列に既に値がある行は触らない(手動修正が常に優先)
//   ・出荷済み(消込台帳)を最優先で割り当て(物理的に先に取っていった分)。ただし台帳の入荷日=箱の到着日の箱だけ
//   ・現役の注文は「注文日≦発注日」のものだけ確定(後から来た注文は引き当て時のFIFOに任せる)
//   ・全量1注文=番号だけ / 分割=「番号:個数」カンマ区切り(セルは薄黄で要確認マーク)

// ♻️ P列を書き直す: EMSリストの「到着済」行のP列を一旦すべて消して、現行ロジックで自動記入し直す。
// バグ時代の古い書き込み(残骸=幽霊の名指し・誤った個数配分)を一括で除去するためのもの。
// 手で書いた名指しも消える(商品コード末尾の（受注番号）タグは書き直しで再現される)ので確認してから実行。
// 在庫反映済みなど過去の行は触らない(履歴として保持)。
function P列を書き直す(){ 直列_(P列を書き直す本体_); }
function P列を書き直す本体_(){
  const ui=SpreadsheetApp.getUi();
  const ans=ui.alert('P列の書き直し(到着済のみ)',
    'EMSリストの「到着済」行のP列(注文番号)を全部消して、今のロジックで書き直します。\n\n'+
    '・バグ時代の古い割当(残骸)が一掃されます\n'+
    '・手で書いた名指しも消えます(コード末尾の（受注番号）タグは自動で再現)\n'+
    '・在庫反映済みなど過去の行は触りません\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  const r=P列書き直し実行_();
  if(r.error){ ui.alert('書き直しでエラー:\n'+r.error); return; }
  SpreadsheetApp.getActive().toast(
    'P列書き直し: クリア'+r.クリア+'行 → 記入'+r.記入+'行'+(r.分割?'（分割'+r.分割+'行）':'')+
    ' / 在庫扱い'+r.在庫+'行。仕上げに ②引き当て実行 を回してください','♻️P列',8);
}

// P列書き直しの本体(ダイアログなし)。到着済行のP列クリア→自動記入。⚖️便の引き直しからも使う
function P列書き直し実行_(){
  const cfg=P_KAKUTEI_CFG;
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!sh) return {error:'発注共有ファイルに「'+cfg.シート+'」がありません'};
  const hr=cfg.ヘッダー行, last=sh.getLastRow();
  if(last<=hr) return {error:'EMSリストにデータがありません'};
  const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  const cSt=f('ステータス列','ステータス'), cP=f('注文番号');
  if(cP<0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};
  const n=last-hr;
  const st=cSt>=0? sh.getRange(hr+1,cSt+1,n,1).getDisplayValues() : null;
  const pRange=sh.getRange(hr+1,cP+1,n,1);
  const pv=pRange.getDisplayValues();
  const clearedA1=[];
  let cleared=0;
  for(let i=0;i<n;i++){
    if(st && String(st[i][0]||'').trim()!=='到着済') continue; // 到着済だけ(ステータス列が無ければ全行)
    if(String(pv[i][0]||'').trim()===''){ continue; }
    pv[i][0]=''; cleared++;
    clearedA1.push(sh.getRange(hr+1+i, cP+1).getA1Notation());
  }
  pRange.setValues(pv);
  if(clearedA1.length) sh.getRangeList(clearedA1).setBackground(null); // 薄黄の要確認マークもクリア
  SpreadsheetApp.flush();
  const r=発注共有P列記入_();
  if(r.error) return {error:r.error};
  return {クリア:cleared, 記入:r.記入, 分割:r.分割, 在庫:r.在庫};
}

// ⚖️ 便の引当をやり直す: 指定した到着日の入荷日を白紙に→P列書き直し→②まで一括実行。
// バグ時代に「新しい注文が先に確定」してしまった便を、古い注文順で公平に引き直すためのリセット。
// 台湾/中国ルートの行は触らない。手入力の同日付スタンプも消える(②が正しい順で貼り直す)。
function 便の引当をやり直す(){ 直列_(便の引当をやり直す本体_); }
function 便の引当をやり直す本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  const M=列マップ_(recv);
  if(M.入荷<0){ ui.alert('受注明細に「入荷日」列がありません'); return; }
  const resp=ui.prompt('⚖️ 便の引当をやり直す',
    'やり直す便の到着日を入力してください(例 2026-07-09)。\n\n'+
    'この日付の入荷日(取り寄せ行・台湾/中国ルートは除く)を全部白紙にし、\n'+
    'P列を書き直して、②引き当て実行まで自動で回します。\n'+
    '→ 古い注文から順に引き当て直されます。',
    ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const d=ymd_(String(resp.getResponseText()||'').trim());
  if(!d || !/^\d{4}-\d{2}-\d{2}$/.test(d)){ ui.alert('日付が読めません: '+resp.getResponseText()); return; }
  const R=recv.getDataRange().getValues();
  const a1=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    if(ymd_(row[M.入荷])!==d) continue;
    const 別ルート=/台湾|中国/.test(String(row[M.選択肢]||'')) || (M.商品名>=0 && /台湾|中国/.test(String(row[M.商品名]||'')));
    if(別ルート) continue;
    a1.push(recv.getRange(i+1, M.入荷+1).getA1Notation());
  }
  if(!a1.length){ ui.alert('入荷日が '+d+' の取り寄せ行はありません。'); return; }
  const ans=ui.alert('確認',
    d+' の入荷日 '+a1.length+' 行を白紙にして、古い注文順で引き直します。\n(P列書き直し→②引き当て実行まで自動で回ります) ええ？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  recv.getRangeList(a1).clearContent();
  SpreadsheetApp.flush();
  const r=P列書き直し実行_();
  if(r.error){ ui.alert('P列書き直しでエラー:\n'+r.error); return; }
  引当実行_本体_(); // ②まで一気に。完了ダイアログの突合せで結果を確認
}

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
  try{ ems=発注共有を開く_().getSheetByName(cfg.シート); }
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
  // ただし付けてよいのは「その注文が実際に受けた箱」=台帳の入荷日と箱の到着日が一致する行だけ。
  // (日付を見ないと、直近に出荷された注文が後から到着した別の箱まで吸い付き、
  //  在庫買い分の箱に発送済みの古い受注番号が大量に記入される)
  消込台帳_出荷済み行_().forEach(t=>{
    const keys=キー展開_(t.sku, t.code); if(!keys.size) return;
    lines.push({ban:t.ban, keys, need:t.qty, date:new Date(0), seq:lines.length, shipped:true, 入荷ymd:ymd_(t.入荷日)});
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
  const colA=(eh.indexOf('EMS到着日')>=0? eh.indexOf('EMS到着日') : (eh.indexOf('到着日')>=0? eh.indexOf('到着日') : 4))+1; // E列既定
  const colE=eh.indexOf('EMS番号')+1;
  if(colP===0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};

  const n=last-hr;
  const pu=ems.getRange(hr+1,colPu,n,1).getDisplayValues();
  const cd=ems.getRange(hr+1,colC,n,1).getDisplayValues();
  const qy=ems.getRange(hr+1,colQ,n,1).getValues();
  const pv=ems.getRange(hr+1,colP,n,1).getDisplayValues();
  const av=ems.getRange(hr+1,colA,n,1).getDisplayValues();
  const ev=colE>0? ems.getRange(hr+1,colE,n,1).getDisplayValues() : null;
  const rows=[];
  for(let i=0;i<n;i++){
    if(!ev || !実EMS番号_(ev[i][0])) continue; // 実EMS番号が無い行・棚卸箱へP列を書かない
    const pno=String(pu[i][0]||'').trim(), code=String(cd[i][0]||'').trim();
    if(!pno||!code) continue;
    const m=pno.match(/^(\d{4})(\d{2})(\d{2})/); if(!m) continue; // 購入No.先頭8桁=発注日
    rows.push({ i, 末:new Date(+m[1],+m[2]-1,+m[3],23,59,59),
      keys:codeKeys_(code), 号:月号_(normCode_(code)), qty:Number(qy[i][0])||0, p:String(pv[i][0]||'').trim(),
      着:ymd_(av[i][0]),        // 箱の到着日(yyyy-MM-dd)。出荷済みの紐付け判定に使う
      tag:タグ受注番号_(code) }); // コード末尾の（受注番号）タグ=人の名指し(中古品の名指し買付など)
  }
  // 受注番号→注文行(タグの名指しで必要数を消し込むためのインデックス)
  const byBan={};
  lines.forEach(l=>(byBan[l.ban]=byBan[l.ban]||[]).push(l));

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
  const target=rows.filter(r=>!r.p && r.qty>0 && (r.tag || r.末.getTime()>=受注最古)).sort((a,b)=>a.末-b.末||a.i-b.i);
  let 記入=0, 分割=0, 在庫=0; const writes=[];
  target.forEach(r=>{
    // コードに（受注番号）タグがある行=人の名指し。日付や照合を問わずその受注番号をそのまま記入(全量)。
    // 同じ受注の必要数は消し込んで、他の箱への二重割当を防ぐ。
    if(r.tag){
      let left=r.qty;
      for(const l of (byBan[r.tag]||[])){ if(left<=0) break; if(l.need<=0) continue;
        const t=Math.min(left,l.need); l.need-=t; left-=t; }
      writes.push({ i:r.i, text:r.tag, warn:false });
      記入++; return;
    }
    const seen=new Set(), cand=[];
    r.keys.forEach(k=>(byKey[k]||[]).forEach(l=>{ if(!seen.has(l.seq)){ seen.add(l.seq); cand.push(l); } }));
    cand.sort(順_);
    let left=r.qty; const got=[];
    for(const l of cand){
      if(left<=0) break;
      if(l.need<=0) continue;
      if(l.shipped && (!l.入荷ymd || !r.着 || l.入荷ymd!==r.着)) continue; // 出荷済みは「実際に受けた箱(入荷日=到着日)」だけ
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
