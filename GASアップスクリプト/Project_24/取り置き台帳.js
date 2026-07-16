const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き初期登録', 要確認:'取り置き要確認', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ'],
  初期HDR:['取置ID','受注番号','商品コード','SKU','注文数量','現物取り置き数量','メモ','判定'],
  要確認HDR:['取置ID','受注番号','商品コード','理由'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});
const 戻しHDR=['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ'];
const Yahoo候補HDR=['取置ID','商品コード','数量','元EMS番号','処理ID','確認'];

function CSV行を受注行オブジェクトへ_(header, rows){
  const head=(header||[]).map(v=>String(v||'').trim()), index=name=>head.indexOf(name);
  const cBan=index('受注番号'), cStatus=index('受注ステータス'), cCode=index('商品コード'), cQty=index('個数');
  const cSku=index('商品SKU')>=0?index('商品SKU'):index('SKU');
  const missing=[];
  if(cBan<0) missing.push('受注番号'); if(cStatus<0) missing.push('受注ステータス');
  if(cCode<0) missing.push('商品コード'); if(cQty<0) missing.push('個数'); if(cSku<0) missing.push('商品SKU/SKU');
  if(missing.length) throw new Error('全ステータスCSVの見出し不足: '+missing.join(','));
  return (rows||[]).map(row=>({受注番号:String(row[cBan]||'').replace(/^niyantarose-/i,''),受注ステータス:String(row[cStatus]||''),
    商品コード:String(row[cCode]||''),SKU:String(row[cSku]||''),個数:Number(row[cQty])||0}));
}

function 取り置き_CSV遷移計画_(csvRows, ledgerRows, now){
  const statuses={}, errors=[], review=[];
  (csvRows||[]).forEach(r=>{
    const ban=String(r.受注番号||'').replace(/^niyantarose-/i,'').trim();
    const key=取り置き_行キー_({ban,code:r.商品コード,sku:r.SKU});
    if(!ban || !取り置き_商品コード_(r.SKU,r.商品コード)) return;
    const next=/キャンセル/.test(String(r.受注ステータス||''))?'キャンセル':/処理済|発送済|出荷済/.test(String(r.受注ステータス||''))?'発送済み':'継続';
    if(statuses[key] && statuses[key]!==next) errors.push('同じ受注行にステータス競合: '+key);
    statuses[key]=next;
  });
  if(errors.length) return {rows:[],review,errors};
  const rows=(ledgerRows||[]).map(r=>{
    const copy=Object.assign({},r); if(copy.状態!==TORIOKI_STATUS.ACTIVE) return copy;
    const key=取り置き_行キー_(copy), next=statuses[key];
    if(next==='発送済み'){ copy.状態=TORIOKI_STATUS.SHIPPED; copy.更新日時=now; }
    else if(next==='キャンセル'){ copy.状態=TORIOKI_STATUS.RETURN; copy.戻し処理結果=TORIOKI_RETURN.UNCHECKED; copy.更新日時=now; }
    else if(!next) review.push({取置ID:copy.取置ID,受注番号:copy.受注番号,商品コード:copy.商品コード,理由:'最新CSVに受注行なし'});
    return copy;
  });
  return {rows,review,errors:[]};
}

function 取り置き_CSV遷移を反映_(plan){
  if(plan.errors && plan.errors.length) throw new Error(plan.errors.join('\n'));
  取り置き台帳_保存_(plan.rows||[]);
  取り置き_表を保存_(TORIOKI_CFG.要確認,TORIOKI_CFG.要確認HDR,plan.review||[]);
}

// 第2引数以降に受注番号Setを何個でも渡せる(部分在庫・希望日待ち・出荷GO未入金・出荷可能)。
// 未入金滞留中に同一商品の新箱が来ても二重引当にならないよう、棚に現物がある注文は全て候補にする。
function 取り置き_初期候補_(orders){
  const banSets=Array.prototype.slice.call(arguments,1).filter(s=>s && typeof s.has==='function');
  return (orders||[]).filter(o=>banSets.some(s=>s.has(String(o.ban)))).map(o=>{
    const key=取り置き_行キー_(o);
    return {取置ID:'INIT|'+key,受注番号:String(o.ban),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),注文数量:Number(o.qty)||0,現物取り置き数量:'',メモ:'',判定:''};
  });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now){
  const errors=[], targets={}, inputIds=new Set();
  // 発送済み・解除済みになった開始前在庫行は履歴。再確定で復活も消滅もさせない(誤操作ガード)
  const lockedIds=new Set((existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態!==TORIOKI_STATUS.ACTIVE).map(r=>String(r.取置ID||'')));
  (inputRows||[]).forEach((r,index)=>{
    inputIds.add(String(r.取置ID||''));
    const raw=r.現物取り置き数量, blank=raw==null || raw==='', entered=blank?0:Number(raw), ordered=Number(r.注文数量)||0;
    if(!blank && (!String(raw).trim() || !Number.isFinite(entered))) errors.push('初期登録'+(index+2)+'行 / 受注'+r.受注番号+': 数量は数値で入力');
    else if(entered<0 || !Number.isInteger(entered)) errors.push('初期登録'+(index+2)+'行: 数量は0以上の整数');
    if(entered>ordered) errors.push('受注'+r.受注番号+': 現物'+entered+'が注文'+ordered+'を超過');
    if(entered>0 && lockedIds.has(String(r.取置ID||''))) errors.push('受注'+r.受注番号+': 既に発送済み/解除済みの初期登録行は変更できません');
    if(entered>0) targets[r.取置ID]=Object.assign({},r,{取り置き数量:entered});
  });
  if(errors.length) return {rows:[],errors};
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'','終了理由・メモ':String(r.メモ||'')});
  });
  return {rows:kept,errors:[]};
}

function 取り置き_戻し確認計画_(inputs, ledger, now){
  const byId={}; (inputs||[]).forEach(r=>byId[String(r.取置ID||'')]=String(r.現物確認||'').trim());
  const errors=[];
  Object.keys(byId).forEach(id=>{ if(byId[id] && byId[id]!=='現物あり' && byId[id]!=='在庫なし') errors.push(id+': 現物確認は「現物あり」か「在庫なし」'); });
  if(errors.length) return {rows:[],errors};
  return {rows:(ledger||[]).map(r=>{
    const choice=byId[String(r.取置ID||'')]; if(!choice) return Object.assign({},r);
    if(r.状態!=='キャンセル戻し' || r.戻し処理結果!=='未確認'){ errors.push(r.取置ID+': 未確認のキャンセル戻しではない'); return Object.assign({},r); }
    return Object.assign({},r,{戻し処理結果:choice,更新日時:now});
  }),errors};
}

function EMS在庫移動_追加計画_(candidates, existing, now){
  const byId={}; (existing||[]).forEach(r=>{ byId[String(r.処理ID||'')]=r; });
  const added=[], errors=[];
  (candidates||[]).forEach(c=>{
    const id=String(c.処理ID||'');
    if(!c.処理ID || !取り置き_整数_(c.数量)) errors.push('Yahoo移動の処理IDまたは数量が不正: '+id);
    else if(byId[id]){
      // 同じ処理IDで数量が変わった=再締めで余りが増減した記録漏れの兆候。黙って捨てず締めを止める
      if(取り置き_整数_(byId[id].数量)!==取り置き_整数_(c.数量))
        errors.push('同じ処理IDで数量が一致しません(便の再締め/記録漏れの疑い): '+id+' 記録'+byId[id].数量+'／今回'+c.数量);
    }
    else {
      const row=Object.assign({},c,{移動先:'Yahoo即納',処理日時:now});
      byId[id]=row; added.push(row);
    }
  });
  return {rows:(existing||[]).map(r=>Object.assign({},r)).concat(added),added,errors};
}

function EMS在庫移動_箱計画_(surplus, existing, now){
  const candidates=(surplus||[]).filter(s=>s.qty>0).map(s=>{
    const ems=String(s.ems||'').trim(), sourceCode=String(s.sourceCode||s.code||'').trim();
    const sourceIdentity=取り置き_供給キー_('',sourceCode).slice(1);
    return {
      処理ID:ems&&sourceIdentity?'YAHOO|EMS|'+ems+'|'+sourceIdentity:'',
      EMS番号:ems,商品コード:sourceCode,数量:s.qty
    };
  });
  return EMS在庫移動_追加計画_(candidates,existing,now);
}

function EMS在庫移動_戻し計画_(returns, existing, now){
  const candidates=(returns||[])
    .filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり')
    .map(r=>({
      処理ID:r.取置ID?'YAHOO|RETURN|'+r.取置ID:'',
      EMS番号:String(r.元EMS番号||''),
      商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),
      数量:r.取り置き数量
    }));
  return EMS在庫移動_追加計画_(candidates,existing,now);
}

function 取り置き_表を読む_(sheetName, headers){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2) return [];
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const index={}; headers.forEach(h=>index[h]=head.indexOf(h));
  if(headers.some(h=>index[h]<0)) throw new Error(sheetName+'の見出し不足: '+headers.filter(h=>index[h]<0).join(','));
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row=>{
    const obj={}; headers.forEach(h=>obj[h]=row[index[h]]); return obj;
  }).filter(obj=>String(obj[headers[0]]||'').trim());
}

function 取り置き_表を保存_(sheetName, headers, rows){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName); if(!sh) sh=ss.insertSheet(sheetName);
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+1) sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const dataRows=Math.max(rows.length,Math.max(0,sh.getLastRow()-1)), dataCols=headers.length;
  if(dataRows>0){
    const values=Array.from({length:dataRows},(_,rowIndex)=>{
      const source=rows[rowIndex];
      return headers.map(header=>source&&source[header]!=null?source[header]:'');
    });
    sh.getRange(2,1,dataRows,dataCols).setValues(values);
  }
}

function 取り置き台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR); }
function 取り置き台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR,rows); }
function EMS在庫移動台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR); }
function EMS在庫移動台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR,rows); }

function 取り置き台帳_割当計画後行_(plan,existingRows,now){
  const rows=(existingRows||[]).map(r=>Object.assign({},r)), byId={};
  rows.forEach((r,index)=>byId[String(r.取置ID||'')]=index);
  (plan&&plan.returnUpdates||[]).forEach(update=>{
    const id=String(update.取置ID||''),index=byId[id]; if(index===undefined) return;
    rows[index]=Object.assign({},rows[index],update,{登録日時:rows[index].登録日時,更新日時:now});
  });
  (plan&&plan.newRows||[]).forEach(row=>{
    const id=String(row.取置ID||''),index=byId[id];
    if(index===undefined){
      const added=Object.assign({},row,{登録日時:now,更新日時:now}); byId[id]=rows.length; rows.push(added);
    }else{
      rows[index]=Object.assign({},rows[index],row,{登録日時:rows[index].登録日時,更新日時:now});
    }
  });
  return rows;
}

function 取り置き台帳_割当計画を反映_(plan,existingRows,now){
  const rows=取り置き台帳_割当計画後行_(plan,existingRows,now);
  取り置き台帳_保存_(rows);
  return rows;
}

function 取り置き_受注番号集合_(sheetName){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName), out=new Set();
  if(!sh || sh.getLastRow()<2) return out;
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const col=head.indexOf('受注番号'); if(col<0) throw new Error(sheetName+'に受注番号見出しがありません');
  sh.getRange(2,col+1,sh.getLastRow()-1,1).getDisplayValues().forEach(r=>{ const ban=String(r[0]||'').trim(); if(ban) out.add(ban); });
  return out;
}

// 書き込み系ボタンは全て直列_(DocumentLock)で排他する(①〜④と同じ約束事)。
// 内部からの再帰呼び出しは本体_を使う(DocumentLockは再入不可のため)。
function 取り置き初期登録を作成(){ 直列_(取り置き初期登録を作成本体_); }
function 取り置き初期登録を作成本体_(){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注), ui=SpreadsheetApp.getUi();
  if(!recv){ ui.alert('受注明細がありません'); return; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), orders=[];
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0 || 区分_(row[M.選択肢])!=='取り寄せ') continue;
    if(引当_別ルート判定_(row[M.選択肢], M.商品名>=0?row[M.商品名]:'')) continue; // 台湾・中国は台帳の対象外
    orders.push({ban,code:String(row[M.コード]||''),sku:M.SKU>=0?String(row[M.SKU]||''):'',qty});
  }
  // 出荷GO未入金(未入金滞留)・出荷可能(発送前)も棚に現物がある=候補に含める(空欄なら未取り置き扱い)
  const candidates=取り置き_初期候補_(orders,
    取り置き_受注番号集合_(HIKIATE_CFG.部分),取り置き_受注番号集合_(HIKIATE_CFG.希望),
    取り置き_受注番号集合_(HIKIATE_CFG.取置),取り置き_受注番号集合_(HIKIATE_CFG.出荷));
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates);
  ui.alert('取り置き初期登録を作成しました','候補'+candidates.length+'行です(部分在庫・希望日待ち・出荷GO未入金・出荷可能)。\n棚の現物を確認し「現物取り置き数量」だけ入力してください(棚に無い行は空欄のまま)。',ui.ButtonSet.OK);
}

function 取り置き初期登録を確定(){ 直列_(取り置き初期登録を確定本体_); }
function 取り置き初期登録を確定本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  const plan=取り置き_初期確定計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('初期登録を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const selected=inputs.filter(r=>(Number(r.現物取り置き数量)||0)>0), qty=selected.reduce((s,r)=>s+(Number(r.現物取り置き数量)||0),0);
  const answer=ui.alert('初期取り置きを確定します','対象'+selected.length+'行 / 合計'+qty+'個です。確定しますか？',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows);
  SpreadsheetApp.getActive().toast('初期取り置き '+selected.length+'行 / '+qty+'個を確定しました','取り置き台帳',7);
}

function キャンセル戻し確認を更新(){ 直列_(キャンセル戻し確認を更新本体_); }
function キャンセル戻し確認を更新本体_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='未確認').map(r=>({
    取置ID:r.取置ID,受注番号:r.受注番号,商品コード:r.商品コード,数量:r.取り置き数量,元EMS番号:r.元EMS番号,現物確認:'',メモ:r['終了理由・メモ']||''
  }));
  取り置き_表を保存_(TORIOKI_CFG.戻し,戻しHDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.戻し);
  if(rows.length) sh.getRange(2,6,rows.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['現物あり','在庫なし'],true).setAllowInvalid(false).build());
  SpreadsheetApp.getActive().toast('未確認のキャンセル戻し '+rows.length+'件','取り置き台帳',6);
}

function キャンセル戻し確認を確定(){ 直列_(キャンセル戻し確認を確定本体_); }
function キャンセル戻し確認を確定本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.戻し,戻しHDR);
  const plan=取り置き_戻し確認計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('キャンセル戻しを確定できません',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const answer=ui.alert('現物確認を確定します','入力済み'+inputs.filter(r=>r.現物確認).length+'件を台帳へ反映します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows); キャンセル戻し確認を更新本体_(); Yahoo戻し候補を更新_();
}

function Yahoo戻し候補を更新_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり').map(r=>({
    取置ID:r.取置ID,商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),数量:r.取り置き数量,元EMS番号:r.元EMS番号,処理ID:'YAHOO|RETURN|'+r.取置ID,確認:''
  }));
  取り置き_表を保存_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.Yahoo候補);
  if(rows.length) sh.getRange(2,6,rows.length,1).insertCheckboxes();
}

function キャンセル戻しをYahoo反映済みにする(){ 直列_(キャンセル戻しをYahoo反映済みにする本体_); }
function キャンセル戻しをYahoo反映済みにする本体_(){
  const ui=SpreadsheetApp.getUi(), candidates=取り置き_表を読む_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR).filter(r=>r.確認===true||String(r.確認).toUpperCase()==='TRUE');
  if(!candidates.length){ ui.alert('Yahoo戻し候補にチェックがありません'); return; }
  const answer=ui.alert('Yahoo戻しの最終確認','Yahoo在庫へ実際に加算済みの'+candidates.length+'件だけを確定します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  const checkedIdCounts={};
  candidates.forEach(r=>{ const id=String(r.取置ID||''); checkedIdCounts[id]=(checkedIdCounts[id]||0)+1; });
  const duplicateChecked=Object.keys(checkedIdCounts).filter(id=>!id || checkedIdCounts[id]!==1);
  if(duplicateChecked.length){ ui.alert('Yahoo戻し候補が重複またはIDなしのため中止しました',duplicateChecked.join('\n'),ui.ButtonSet.OK); return; }
  const ledger=取り置き台帳_読む_(), selectedIds=new Set(candidates.map(r=>String(r.取置ID)));
  const ledgerIdCounts={};
  ledger.forEach(r=>{ const id=String(r.取置ID||''); ledgerIdCounts[id]=(ledgerIdCounts[id]||0)+1; });
  const unresolved=Array.from(selectedIds).filter(id=>!id || ledgerIdCounts[id]!==1);
  if(unresolved.length){ ui.alert('Yahoo戻し候補が現在の取り置き台帳と一致しないため中止しました',unresolved.join('\n'),ui.ButtonSet.OK); return; }
  const returns=ledger.filter(r=>selectedIds.has(String(r.取置ID)));
  const candidateById={}; candidates.forEach(r=>candidateById[String(r.取置ID)]=r);
  const tampered=returns.filter(r=>{
    const id=String(r.取置ID), candidate=candidateById[id];
    const sourceCode=String(r.元EMS商品コード||r.商品コード||'').trim();
    return String(candidate.商品コード||'').trim()!==sourceCode ||
      取り置き_整数_(candidate.数量)!==取り置き_整数_(r.取り置き数量) ||
      String(candidate.元EMS番号||'').trim()!==String(r.元EMS番号||'').trim() ||
      String(candidate.処理ID||'').trim()!=='YAHOO|RETURN|'+id;
  });
  if(tampered.length){ ui.alert('Yahoo戻し候補が現在の取り置き台帳と一致しないため中止しました',tampered.map(r=>r.取置ID).join('\n'),ui.ButtonSet.OK); return; }
  if(returns.some(r=>r.状態!=='キャンセル戻し'||r.戻し処理結果!=='現物あり')){ ui.alert('現物ありでない候補が含まれるため中止しました'); return; }
  const existingMoves=EMS在庫移動台帳_読む_(), plan=EMS在庫移動_戻し計画_(returns,existingMoves,new Date());
  if(plan.errors.length){ ui.alert('Yahoo移動を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const updatedLedger=ledger.map(r=>selectedIds.has(String(r.取置ID))?Object.assign({},r,{戻し処理結果:'Yahoo反映済み',更新日時:new Date()}):r);
  try{ EMS在庫移動台帳_保存_(plan.rows); }
  catch(error){ ui.alert('Yahoo移動台帳の保存に失敗しました',error.message,ui.ButtonSet.OK); return; }
  try{ 取り置き台帳_保存_(updatedLedger); }
  catch(ledgerError){
    try{ EMS在庫移動台帳_保存_(existingMoves); }
    catch(rollbackError){
      ui.alert('取り置き台帳の保存とYahoo移動台帳の復旧に失敗しました',
        '取り置き台帳: '+ledgerError.message+'\nYahoo移動台帳: '+rollbackError.message,ui.ButtonSet.OK);
      return;
    }
    ui.alert('取り置き台帳の保存に失敗したためYahoo移動台帳を元へ戻しました',ledgerError.message,ui.ButtonSet.OK);
    return;
  }
  Yahoo戻し候補を更新_();
}

function 選択した取り置きを手動解除(){ 直列_(選択した取り置きを手動解除本体_); }
function 選択した取り置きを手動解除本体_(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getActiveSheet(), ui=SpreadsheetApp.getUi(), row=sh.getActiveRange().getRow();
  if(sh.getName()!==TORIOKI_CFG.台帳 || row<2){ ui.alert('取り置き台帳の解除する行を選択してください'); return; }
  // 行番号ではなく選択行の取置IDで台帳の対象を特定する(空行・手動の行操作でズレても別の行を解除しない)
  const selectedId=String(sh.getRange(row,1).getDisplayValue()||'').trim();
  if(!selectedId){ ui.alert('選択した行に取置IDがありません'); return; }
  const ledger=取り置き台帳_読む_(), matches=ledger.filter(r=>String(r.取置ID||'')===selectedId);
  if(matches.length!==1){ ui.alert('取置IDが台帳で一意に特定できません: '+selectedId); return; }
  const target=matches[0];
  if(target.状態!=='取り置き中'){ ui.alert('取り置き中の行だけ手動解除できます'); return; }
  const response=ui.prompt('手動解除の理由','登録間違い、現物不足などの理由を入力してください。',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK) return;
  const reason=String(response.getResponseText()||'').trim(); if(!reason){ ui.alert('解除理由は必須です'); return; }
  取り置き台帳_保存_(ledger.map(r=>String(r.取置ID||'')===selectedId
    ? Object.assign({},r,{状態:'手動解除','終了理由・メモ':reason,更新日時:new Date()}) : r));
}
