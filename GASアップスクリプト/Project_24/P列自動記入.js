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
// 到着日は1日でも、カンマ区切りで複数日(例 2026-07-09, 2026-07-10)でも指定可。
function 便の引当をやり直す(){ 直列_(便の引当をやり直す本体_); }
function 便の引当をやり直す_日付解析_(text){
  const raw=String(text||'').trim(); if(!raw) return {error:'日付が空です'};
  const dates=[], bad=[];
  raw.split(/[,、\s]+/).map(s=>s.trim()).filter(Boolean).forEach(tok=>{
    const d=ymd_(tok);
    if(!d || !/^20\d{2}-\d{2}-\d{2}$/.test(d)) bad.push(tok);
    else if(dates.indexOf(d)<0) dates.push(d);
  });
  if(bad.length) return {error:'日付が読めません: '+bad.join(', ')};
  if(!dates.length) return {error:'有効な到着日がありません'};
  dates.sort();
  return {dates:dates};
}
function 便の引当をやり直す本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  const M=列マップ_(recv);
  if(M.入荷<0){ ui.alert('受注明細に「入荷日」列がありません'); return; }
  const resp=ui.prompt('⚖️ 便の引当をやり直す',
    'やり直す便の到着日を入力してください。\n'+
    '例: 2026-07-09\n'+
    '複数便: 2026-07-09, 2026-07-10（カンマ・スペース区切り）\n\n'+
    '指定した日付の入荷日(取り寄せ行・台湾/中国ルートは除く)を全部白紙にし、\n'+
    'P列を書き直して、②引き当て実行まで自動で回します。\n'+
    '→ 古い注文から順に引き当て直されます。',
    ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const parsed=便の引当をやり直す_日付解析_(resp.getResponseText());
  if(parsed.error){ ui.alert('日付が読めません:\n'+parsed.error); return; }
  const dateSet={}; parsed.dates.forEach(d=>{ dateSet[d]=1; });
  const R=recv.getDataRange().getValues();
  const a1=[], perDate={};
  parsed.dates.forEach(d=>{ perDate[d]=0; });
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const d=ymd_(row[M.入荷]); if(!dateSet[d]) continue;
    const 別ルート=/台湾|中国/.test(String(row[M.選択肢]||'')) || (M.商品名>=0 && /台湾|中国/.test(String(row[M.商品名]||'')));
    if(別ルート) continue;
    a1.push(recv.getRange(i+1, M.入荷+1).getA1Notation());
    perDate[d]++;
  }
  if(!a1.length){ ui.alert('入荷日が '+parsed.dates.join(' / ')+' の取り寄せ行はありません。'); return; }
  const 内訳=parsed.dates.map(d=> d+' '+perDate[d]+'行').join('、');
  const ans=ui.alert('確認',
    parsed.dates.length+'日分・計 '+a1.length+' 行の入荷日を白紙にして、古い注文順で引き直します。\n'+
    '（'+内訳+'）\n\nP列書き直し→②引き当て実行まで自動で回ります。 ええ？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  recv.getRangeList(a1).clearContent();
  SpreadsheetApp.flush();
  const r=P列書き直し実行_();
  if(r.error){ ui.alert('P列書き直しでエラー:\n'+r.error); return; }
  引当実行_本体_(); // ②まで一気に。完了ダイアログの突合せで結果を確認
}

function P列処理対象EMS_(status){
  return String(status==null?'':status).trim()==='到着済';
}

// 実EMSで実際に到着した日付を商品キーごとに保持する。
// 「在庫反映済み」も含めることで、履歴導入前の旧便を新しい便へ付け替えないための根拠にする。
function EMS到着実績Map_(rows){
  const map={};
  (rows||[]).forEach(r=>{
    const status=String(r&&r.status||'').trim(), d=ymd_(r&&r.着);
    if(!d || (status!=='到着済' && status!=='在庫反映済み')) return;
    (r.keys||[]).forEach(k=>{
      const key=String(k||'').trim(); if(!key) return;
      (map[key]=map[key]||new Set()).add(d);
    });
  });
  return map;
}

// 実EMSの「商品キー × 到着日」ごとの便番号を保持する。
// 入荷日が旧便を指しているのに受注明細のEMS番号だけ今回便のまま残った時の訂正根拠に使う。
function EMS到着便Map_(rows){
  const map={};
  (rows||[]).forEach(r=>{
    const status=String(r&&r.status||'').trim(), d=ymd_(r&&r.着), ems=String(r&&r.ems||'').trim();
    if(!d || !実EMS番号_(ems) || (status!=='到着済' && status!=='在庫反映済み')) return;
    (r.keys||[]).forEach(k=>{
      const key=String(k||'').trim(); if(!key) return;
      const byDate=map[key]||(map[key]={});
      (byDate[d]=byDate[d]||new Set()).add(ems);
    });
  });
  return map;
}

// 同じ商品候補・同じ入荷日に実在するEMS番号を返す。
// 該当便が無い場合は null（従来の保持ロジックへフォールバック）、該当時は重複を除いた配列。
function 入荷日実EMS候補_(keys, arrival, map){
  const d=ymd_(arrival); if(!d || !map) return null;
  const out=new Set(); let found=false;
  (keys||[]).forEach(k=>{
    const byDate=map[String(k||'').trim()], set=byDate&&byDate[d];
    if(!set || typeof set.forEach!=='function') return;
    found=true;
    set.forEach(v=>{ if(実EMS番号_(v)) out.add(String(v).trim()); });
  });
  return found?Array.from(out):null;
}

function P列到着日一致_(orderArrival, emsArrival){
  const order=ymd_(orderArrival), ems=ymd_(emsArrival);
  if(!order) return true;
  return !!ems && order===ems;
}

function P列指定解析_(text, rowQty){
  const src=String(text==null?'':text).trim();
  if(!src) return {entries:[],invalid:false};
  const entries=[]; let invalid=false;
  src.split(/[,、]/).forEach(part=>{
    const m=String(part).trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
    if(!m){ invalid=true; return; }
    entries.push({
      ban:m[1],
      qty:m[2]?Number(m[2]):Math.max(0,Number(rowQty)||0),
      explicit:!!m[2]
    });
  });
  return {entries,invalid};
}

function P列指定文字列_(entries, rowQty){
  const a=(entries||[]).filter(e=>e && e.ban && Number(e.qty)>0);
  if(a.length===1 && Number(a[0].qty)===Number(rowQty) && !a[0].explicit) return a[0].ban;
  return a.map(e=>e.ban+':'+e.qty).join(', ');
}

function P列既存指定を再構成_(text, rowQty, consume){
  const parsed=P列指定解析_(text,rowQty), q=Math.max(0,Number(rowQty)||0);
  if(parsed.invalid){
    let consumed=0;
    parsed.entries.forEach(e=>{
      const request=Math.max(0,Math.min(Number(e.qty)||0,q-consumed));
      const accepted=Math.max(0,Math.min(request,Number(consume(e.ban,request))||0));
      consumed+=accepted;
    });
    return {
      text:String(text),entries:[],keptQty:q,removedQty:0,
      blockedBans:Array.from(new Set(parsed.entries.map(e=>e.ban))),invalid:true
    };
  }
  const kept=[]; let keptQty=0, removedQty=0;
  parsed.entries.forEach(e=>{
    const wanted=Math.max(0,Number(e.qty)||0);
    const request=Math.max(0,Math.min(wanted,q-keptQty));
    removedQty+=wanted-request;
    const accepted=Math.max(0,Math.min(request,Number(consume(e.ban,request))||0));
    if(accepted>0) kept.push({ban:e.ban,qty:accepted,explicit:e.explicit});
    keptQty+=accepted; removedQty+=request-accepted;
  });
  return {
    text:P列指定文字列_(kept,q),
    entries:kept,
    keptQty,
    removedQty,
    blockedBans:Array.from(new Set(parsed.entries.map(e=>e.ban))),
    invalid:false
  };
}

function P列FIFO候補_(candidates, blockedBans){
  const blocked=new Set(blockedBans||[]);
  return (candidates||[]).filter(line=>!blocked.has(String(line.ban||'')));
}

// 末尾タグは従来どおり最優先。商品コードそのものが注文番号なら、
// 現役の取り寄せ注文へ一意に結び付く場合だけP列へ同じ受注番号を書く。
function P列名指し受注番号_(code, lines){
  const tagged=タグ受注番号_(code); if(tagged) return tagged;
  const direct=注文番号指定引当先_(code,lines);
  return direct.line?direct.ban:'';
}

function P列に注文番号を自動記入(){
  const r=発注共有P列記入_();
  if(r.error){ SpreadsheetApp.getUi().alert(r.error); return; }
  SpreadsheetApp.getActive().toast(
    'P列自動記入: '+r.記入+'行'+(r.分割?'（分割/一部在庫 '+r.分割+'行=薄黄）':'')+
    ' / 該当注文なし(在庫扱い) '+r.在庫+'行 / 既存 '+r.既存+'行'+
    (r.過剰除外?' / 過剰P除外 '+r.過剰除外+'個':'')+
    (r.解析警告?' / P解析警告 '+r.解析警告+'行':''),'📝P列',8);
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
    const sku=M.SKU>=0?String(row[M.SKU]||'').trim():'';
    const code=String(row[M.コード]||'').trim();
    const keys=キー展開_(sku,code);
    if(!keys.size) continue;
    lines.push({
      ban, code, sku, qty, kbn:'取り寄せ', キャンセル:false,
      keys, need:qty, date, seq:lines.length, shipped:false,
      入荷ymd:M.入荷>=0?ymd_(row[M.入荷]):''
    });
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
  const idxSt=eh.indexOf('ステータス列')>=0?eh.indexOf('ステータス列'):eh.indexOf('ステータス');
  const idxCode=eh.indexOf('商品コード');
  const idxArrival=eh.indexOf('EMS到着日')>=0?eh.indexOf('EMS到着日'):eh.indexOf('到着日');
  const idxEms=eh.indexOf('EMS番号');
  const colSt=idxSt+1;
  const colP=eh.indexOf('注文番号')+1;
  const colPu=(eh.indexOf('購入No.')>=0? eh.indexOf('購入No.'):5)+1;
  const colC=(idxCode>=0? idxCode:8)+1;
  const colQ=(eh.indexOf('数量')>=0? eh.indexOf('数量'):9)+1;
  const colA=(idxArrival>=0? idxArrival:4)+1; // E列既定
  const colE=idxEms+1;
  const 到着実績取得済=idxSt>=0 && idxCode>=0 && idxArrival>=0 && idxEms>=0;
  if(colSt===0) return {error:'EMSリストの'+hr+'行目に「ステータス列」見出しがありません'};
  if(colP===0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};

  const n=last-hr;
  const cols=[colSt,colPu,colC,colQ,colP,colA,colE].filter(c=>c>0);
  const first=Math.min.apply(null,cols), width=Math.max.apply(null,cols)-first+1;
  const block=ems.getRange(hr+1,first,n,width).getValues();
  const at=(row,col)=>col>0?row[col-first]:'';
  const pColumn=block.map(row=>[String(at(row,colP)==null?'':at(row,colP))]);
  const ev=block.map(row=>[at(row,colE)]);
  const rows=[], 到着実績Rows=[];
  for(let i=0;i<n;i++){
    const row=block[i];
    if(!実EMS番号_(ev[i][0])) continue; // 実EMS番号が無い行・棚卸箱へP列を書かない
    const status=String(at(row,colSt)||'').trim(), code=String(at(row,colC)||'').trim();
    if(到着実績取得済 && code) 到着実績Rows.push({status, 着:ymd_(at(row,colA)), keys:codeKeys_(code), ems:String(ev[i][0]||'').trim()});
    const pno=String(at(row,colPu)||'').trim();
    if(!pno||!code) continue;
    const m=pno.match(/^(\d{4})(\d{2})(\d{2})/); if(!m) continue; // 購入No.先頭8桁=発注日
    const pOriginal=String(at(row,colP)==null?'':at(row,colP));
    rows.push({ i, 末:new Date(+m[1],+m[2]-1,+m[3],23,59,59),
      keys:codeKeys_(code), 号:月号_(normCode_(code)), qty:Number(at(row,colQ))||0,
      pOriginal, p:pOriginal.trim(), 対象:P列処理対象EMS_(at(row,colSt)),
      status,
      着:ymd_(at(row,colA)),     // 箱の到着日(yyyy-MM-dd)。注文側の入荷日と照合する
      tag:P列名指し受注番号_(code,lines) }); // 末尾タグ、またはコードそのものが現役受注番号=人の名指し
  }
  const 到着実績=到着実績取得済?EMS到着実績Map_(到着実績Rows):{};
  const 到着便=到着実績取得済?EMS到着便Map_(到着実績Rows):{};
  // 受注番号→注文行(タグの名指しで必要数を消し込むためのインデックス)
  const byBan={};
  lines.forEach(l=>(byBan[l.ban]=byBan[l.ban]||[]).push(l));

  // ---- 1周目: 到着済行の既存Pを、商品と到着日が一致する残需要へ充当する ----
  let 過剰除外=0, 解析警告=0;
  rows.forEach(r=>{
    r.nextP=r.pOriginal;
    if(!r.対象) return; // 締め済み・未着などのPは証跡として変更しない
    if(r.tag){
      let left=r.qty;
      for(const l of (byBan[r.tag]||[])){
        if(left<=0) break;
        if(l.need<=0) continue;
        const take=Math.min(left,l.need); l.need-=take; left-=take;
      }
      r.entries=[{ban:r.tag,qty:r.qty,explicit:false}];
      r.left=0; r.blocked=[r.tag]; r.invalid=false; r.nextP=r.tag;
      return;
    }
    const rebuilt=P列既存指定を再構成_(r.p,r.qty,(ban,qty)=>{
      let left=qty;
      for(const k of r.keys){
        for(const l of (byKey[k]||[])){
          if(left<=0) break;
          if(l.ban!==ban||l.need<=0) continue;
          if(l.shipped && !l.入荷ymd) continue;
          if(!P列到着日一致_(l.入荷ymd,r.着)) continue;
          const take=Math.min(left,l.need); l.need-=take; left-=take;
        }
        if(left<=0) break;
      }
      return qty-left;
    });
    r.entries=rebuilt.entries.slice();
    r.left=Math.max(0,r.qty-rebuilt.keptQty);
    r.blocked=rebuilt.blockedBans;
    r.invalid=rebuilt.invalid;
    r.nextP=rebuilt.text;
    過剰除外+=rebuilt.removedQty;
    if(rebuilt.invalid) 解析警告++;
  });

  // ---- 2周目: 到着済行の残数を、同じ注文以外の未充足注文へFIFOで記入 ----
  const target=rows.filter(r=>r.対象 && !r.tag && !r.invalid && r.left>0 && r.末.getTime()>=受注最古)
    .sort((a,b)=>a.末-b.末||a.i-b.i);
  target.forEach(r=>{
    const seen=new Set(), cand=[];
    r.keys.forEach(k=>(byKey[k]||[]).forEach(l=>{ if(!seen.has(l.seq)){ seen.add(l.seq); cand.push(l); } }));
    const candidates=P列FIFO候補_(cand,r.blocked).sort(順_);
    let left=r.left;
    for(const l of candidates){
      if(left<=0) break;
      if(l.need<=0) continue;
      if(l.shipped && !l.入荷ymd) continue;
      if(!P列到着日一致_(l.入荷ymd,r.着)) continue;
      if(!l.shipped && l.date.getTime()>r.末.getTime()) continue; // 現役の注文だけ「注文日≦発注日」
      if(!l.shipped && r.号 && 期待号_(l.date)!==r.号) continue;   // 月号付き(定期購読)は「注文月+1=その号」の注文だけ
      const take=Math.min(left,l.need); l.need-=take; left-=take;
      const prev=r.entries.find(e=>e.ban===l.ban);
      if(prev) prev.qty+=take;
      else r.entries.push({ban:l.ban,qty:take,explicit:false});
    }
    r.left=left;
    r.nextP=P列指定文字列_(r.entries,r.qty);
  });

  const writes=rows.filter(r=>r.nextP!==r.pOriginal).map(r=>({
    i:r.i,
    text:r.nextP,
    warn:!!r.nextP && (/[:：,、]/.test(r.nextP) || r.left>0)
  }));
  const 記入=writes.filter(w=>w.text).length;
  const 分割=rows.filter(r=>r.対象 && r.nextP && /[:：,、]/.test(r.nextP)).length;
  const 在庫=rows.filter(r=>r.対象 && !r.tag && r.qty>0 && !r.nextP).length;
  if(writes.length){
    writes.forEach(w=>{ pColumn[w.i][0]=w.text; });
    ems.getRange(hr+1,colP,n,1).setValues(pColumn);
    const warn=writes.filter(w=>w.warn).map(w=>ems.getRange(hr+1+w.i,colP).getA1Notation());
    if(warn.length) ems.getRangeList(warn).setBackground('#fff2cc');
  }
  return {記入, 分割, 在庫, 既存:rows.filter(r=>r.p).length, 過剰除外, 解析警告, 到着実績, 到着便, 到着実績取得済};
}
