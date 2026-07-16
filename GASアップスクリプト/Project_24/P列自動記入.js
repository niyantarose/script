// ===== P列自動記入(引当ファイル側から発注共有ファイルのEMSリストへ書き込む) =====
// 受注明細(現役)の残需要と取り置き台帳の供給使用量をEMSリストと照合し、
// メモリ上でP列計画を作ってから一括反映する。
// ②引き当て実行の最初にも自動で走るので、通常は手で押す必要はない。
//
// ルール:
//   ・同じEMS供給の取り置き中は固定表示し、その他の使用済み数量は表示せず供給だけ塞ぐ
//   ・台帳で未使用の残数だけを現役注文の古い順にFIFOで割り当てる
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
  const pv=sh.getRange(hr+1,cP+1,n,1).getDisplayValues();
  const clearedValues=pv.map(row=>[String(!row||row[0]==null?'':row[0])]);
  const clearedRows=[];
  let cleared=0;
  for(let i=0;i<n;i++){
    if(st && String(st[i][0]||'').trim()!=='到着済') continue; // 到着済だけ(ステータス列が無ければ全行)
    if(String(clearedValues[i][0]||'').trim()==='') continue;
    clearedValues[i][0]=''; cleared++; clearedRows.push(i);
  }
  const plan=発注共有P列計画_({currentP:clearedValues});
  if(plan.error) return {error:plan.error};
  const sheetIdentity=P列シート識別子_(sh), planSheetIdentity=P列シート識別子_(plan.sheet);
  if(!sheetIdentity || planSheetIdentity!==sheetIdentity || plan.startRow!==hr+1 || plan.colP!==cP+1 || plan.rowCount!==n ||
      !Array.isArray(plan.values) || plan.values.length!==n){
    return {error:'P列書き直し計画の範囲がEMSリストと一致しません'};
  }
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount,1).setValues(plan.values);
  const backgrounds=clearedRows.map(i=>({a1:P列セルA1_(plan.startRow+i,plan.colP),color:null}))
    .concat(plan.backgrounds||[]);
  P列背景を反映_(plan.sheet, backgrounds);
  return {クリア:cleared, 記入:plan.summary.記入, 分割:plan.summary.分割, 在庫:plan.summary.在庫};
}

// 背景色を色ごとにRangeListへまとめて反映(1セルずつのAPI呼び出しにしない)
function P列背景を反映_(sheet, items){
  const byColor={};
  (items||[]).forEach(item=>{ if(item && item.a1){ const key=String(item.color); (byColor[key]=byColor[key]||[]).push(item.a1); } });
  Object.keys(byColor).forEach(key=>{
    sheet.getRangeList(byColor[key]).setBackground(key==='null'? null : key);
  });
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

function P列指定文字列_(entries, rowQty){
  const a=(entries||[]).filter(e=>e && e.ban && Number(e.qty)>0);
  if(a.length===1 && Number(a[0].qty)===Number(rowQty) && !a[0].explicit) return a[0].ban;
  return a.map(e=>e.ban+':'+e.qty).join(', ');
}

// (旧: P列指定解析_/P列既存指定を再構成_/P列FIFO候補_ は台帳基準化で呼び出し元が無くなり削除。
//  P列の既存値は計画と一致すれば触らず、違えば計画で書き直す=個別の名指しは個別ボタンが台帳へ直接書く)

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

function P列セルA1_(row,col){
  let n=Math.max(1,Number(col)||1), letters='';
  while(n>0){ n--; letters=String.fromCharCode(65+n%26)+letters; n=Math.floor(n/26); }
  return letters+String(row);
}

function P列シート識別子_(sheet){
  try{
    if(!sheet || typeof sheet.getSheetId!=='function' || typeof sheet.getParent!=='function') return '';
    const parent=sheet.getParent();
    if(!parent || typeof parent.getId!=='function') return '';
    const spreadsheetId=String(parent.getId()||'').trim(), sheetId=sheet.getSheetId();
    return spreadsheetId && sheetId!=null?spreadsheetId+'|'+String(sheetId):'';
  }catch(e){ return ''; }
}

function 発注共有P列計画を反映_(plan){
  if(plan.error || !plan.writes || !plan.writes.length) return plan.summary||plan;
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount||plan.values.length,1).setValues(plan.values);
  P列背景を反映_(plan.sheet, plan.backgrounds);
  return Object.assign({},plan.summary,{到着実績:plan.到着実績,到着便:plan.到着便,到着実績取得済:plan.到着実績取得済});
}

function 発注共有P列計画_(options){
  options=options||{};
  const cfg=P_KAKUTEI_CFG, ss=SpreadsheetApp.getActive();
  let ems;
  try{ ems=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!ems) return {error:'発注共有ファイルに「'+cfg.シート+'」がありません'};
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return {error:'「'+HIKIATE_CFG.受注+'」タブがありません'};

  let ledgerSummary;
  // options.追加台帳行: ④が保存前に検知した台帳外出荷の自動登録分。使用済みとして供給から差し引く
  try{ ledgerSummary=取り置き_集計_(取り置き台帳_読む_().concat(options&&options.追加台帳行||[]),EMS在庫移動台帳_読む_()); }
  catch(e){ return {error:'取り置き台帳またはEMS在庫移動台帳が読み込めません:\n'+e.message}; }

  // ---- 現役の取り寄せ行(受注明細) ----
  const M=列マップ_(recv), R=recv.getDataRange().getValues();
  const head=recv.getRange(M.hr,1,1,recv.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const c日時=head.indexOf('注文日時');
  const 日時_=v=>{ if(v instanceof Date) return isNaN(v.getTime())?null:v;
    const s=String(v||'').trim(); if(!s) return null;
    const d=new Date(s.replace(/\//g,'-')); return isNaN(d.getTime())?null:d; };
  const キー展開_=(sku,code)=>new Set(引当用照合キー一覧_(sku,code));
  const lines=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i]; const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    // 台湾・中国ルートで入荷日手入力済みの行は確保済み扱い(韓国EMSのP列で名指ししない)
    if(引当_別ルート判定_(row[M.選択肢], M.商品名>=0?row[M.商品名]:'') &&
       M.入荷>=0 && String(row[M.入荷]==null?'':row[M.入荷]).trim()!=='') continue;
    const qty=Number(row[M.個数])||0; if(qty<=0) continue;
    const date=c日時>=0? 日時_(row[c日時]) : null; if(!date) continue;
    const sku=M.SKU>=0?String(row[M.SKU]||'').trim():'';
    const code=String(row[M.コード]||'').trim();
    const keys=キー展開_(sku,code);
    if(!keys.size) continue;
    const line={
      ban, code, sku, qty, kbn:'取り寄せ', キャンセル:false,
      keys, need:0, date, seq:lines.length
    };
    line.need=取り置き_今回必要数_(line,ledgerSummary);
    lines.push(line);
  }
  if(!lines.length) return {error:'受注明細に取り寄せの注文がありません。先にCSV取込をしてください。'};

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
  const hasCurrentP=Object.prototype.hasOwnProperty.call(options,'currentP');
  if(hasCurrentP && (!Array.isArray(options.currentP) || options.currentP.length!==n)){
    return {error:'P列プレビュー値の行数がEMSリストと一致しません'};
  }
  const pColumn=(hasCurrentP?options.currentP:block.map(row=>[at(row,colP)])).map(value=>{
    const cell=Array.isArray(value)?value[0]:value;
    return [String(cell==null?'':cell)];
  });
  const ev=block.map(row=>[at(row,colE)]);
  const rows=[], 到着実績Rows=[];
  for(let i=0;i<n;i++){
    const row=block[i];
    if(!実EMS番号_(ev[i][0])) continue; // 実EMS番号が無い行・棚卸箱へP列を書かない
    const status=String(at(row,colSt)||'').trim(), code=String(at(row,colC)||'').trim();
    if(到着実績取得済 && code) 到着実績Rows.push({status, 着:ymd_(at(row,colA)), keys:引当用照合キー一覧_('',code), ems:String(ev[i][0]||'').trim()});
    const pno=String(at(row,colPu)||'').trim();
    if(!pno||!code) continue;
    if(!/^\d{8}/.test(pno)) continue;
    const emsNo=String(ev[i][0]||'').trim();
    const pOriginal=pColumn[i][0];
    rows.push({
      i,ems:emsNo,code:normCode_(code),sourceCode:code,directBan:注文番号在庫コード_(code)||タグ受注番号_(code),arrival:ymd_(at(row,colA)),qty:Number(at(row,colQ))||0,
      pOriginal,対象:P列処理対象EMS_(at(row,colSt)),status
    });
  }
  const 到着実績=到着実績取得済?EMS到着実績Map_(到着実績Rows):{};
  const 到着便=到着実績取得済?EMS到着便Map_(到着実績Rows):{};

  const fixedBySupply={};
  Object.keys(ledgerSummary.activeRowsByKey||{}).forEach(orderKey=>{
    (ledgerSummary.activeRowsByKey[orderKey]||[]).forEach(r=>{
      const source=String(r.取置元種別||'').trim(), emsNo=String(r.元EMS番号||'').trim();
      if(r.状態!==TORIOKI_STATUS.ACTIVE || (source!=='EMS' && source!=='キャンセル再引当') || !emsNo) return;
      const key=取り置き_供給キー_(emsNo,r.元EMS商品コード||r.商品コード);
      (fixedBySupply[key]=fixedBySupply[key]||[]).push({ban:String(r.受注番号||'').trim(),qty:Number(r.取り置き数量)||0});
    });
  });
  const calculated=P列計画_純計算_(rows.filter(r=>r.対象),lines,fixedBySupply,ledgerSummary.usageBySupply);
  const writes=[];
  calculated.rows.forEach(r=>{
    if(r.nextP===r.pOriginal) return;
    pColumn[r.i][0]=r.nextP; writes.push(r.i);
  });
  const backgrounds=calculated.rows.filter(r=>writes.indexOf(r.i)>=0 && r.nextP && (/[:：,、]/.test(r.nextP)||r.left>0))
    .map(r=>({a1:P列セルA1_(hr+1+r.i,colP),color:'#fff2cc'}));
  const summary={
    記入:writes.filter(i=>pColumn[i][0]).length,
    分割:calculated.rows.filter(r=>r.nextP && /[:：,、]/.test(r.nextP)).length,
    在庫:calculated.rows.filter(r=>r.qty>0 && !r.nextP).length,
    既存:calculated.rows.filter(r=>String(r.pOriginal||'').trim()).length,
    過剰除外:0,解析警告:0
  };
  return {
    error:'',sheet:ems,startRow:hr+1,colP,rowCount:n,values:pColumn,backgrounds,writes,
    rows:calculated.rows,到着実績,到着便,到着実績取得済,summary
  };
}

function P列計画_確定割当_(plan){
  const out=[];
  (plan.rows||[]).filter(r=>!String(r.directBan||'').trim()).forEach(r=>(r.entries||[]).forEach(e=>{
    out.push({ems:r.ems,code:r.code,sourceCode:String(r.sourceCode||r.code||'').trim(),ban:e.ban,qty:e.qty});
  }));
  return out;
}

function 発注共有P列記入_(){
  const plan=発注共有P列計画_();
  if(plan.error) return {error:plan.error};
  return 発注共有P列計画を反映_(plan);
}
