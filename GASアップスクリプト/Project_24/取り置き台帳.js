const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き初期登録', 要確認:'取り置き要確認', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ'],
  初期HDR:['取置ID','受注番号','商品コード','SKU','注文数量','現物取り置き数量','メモ','判定'],
  要確認HDR:['取置ID','受注番号','商品コード','理由'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});

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

function 取り置き_初期候補_(orders, partialBans, holdBans){
  return (orders||[]).filter(o=>partialBans.has(String(o.ban))||holdBans.has(String(o.ban))).map(o=>{
    const key=取り置き_行キー_(o);
    return {取置ID:'INIT|'+key,受注番号:String(o.ban),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),注文数量:Number(o.qty)||0,現物取り置き数量:'',メモ:'',判定:''};
  });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now){
  const errors=[], targets={}, inputIds=new Set();
  (inputRows||[]).forEach((r,index)=>{
    inputIds.add(String(r.取置ID||''));
    const raw=r.現物取り置き数量, blank=raw==null || raw==='', entered=blank?0:Number(raw), ordered=Number(r.注文数量)||0;
    if(!blank && (!String(raw).trim() || !Number.isFinite(entered))) errors.push('初期登録'+(index+2)+'行 / 受注'+r.受注番号+': 数量は数値で入力');
    else if(entered<0 || !Number.isInteger(entered)) errors.push('初期登録'+(index+2)+'行: 数量は0以上の整数');
    if(entered>ordered) errors.push('受注'+r.受注番号+': 現物'+entered+'が注文'+ordered+'を超過');
    if(entered>0) targets[r.取置ID]=Object.assign({},r,{取り置き数量:entered});
  });
  if(errors.length) return {rows:[],errors};
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'',終了理由・メモ:String(r.メモ||'')});
  });
  return {rows:kept,errors:[]};
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
  if(sh.getMaxRows()>1) sh.getRange(2,1,sh.getMaxRows()-1,sh.getMaxColumns()).clearContent();
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows.map(r=>headers.map(h=>r[h]==null?'':r[h])));
  sh.setFrozenRows(1);
}

function 取り置き台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR); }
function 取り置き台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR,rows); }
function EMS在庫移動台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR); }
function EMS在庫移動台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR,rows); }

function 取り置き_受注番号集合_(sheetName){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName), out=new Set();
  if(!sh || sh.getLastRow()<2) return out;
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const col=head.indexOf('受注番号'); if(col<0) throw new Error(sheetName+'に受注番号見出しがありません');
  sh.getRange(2,col+1,sh.getLastRow()-1,1).getDisplayValues().forEach(r=>{ const ban=String(r[0]||'').trim(); if(ban) out.add(ban); });
  return out;
}

function 取り置き初期登録を作成(){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注), ui=SpreadsheetApp.getUi();
  if(!recv){ ui.alert('受注明細がありません'); return; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), orders=[];
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0 || 区分_(row[M.選択肢])!=='取り寄せ') continue;
    orders.push({ban,code:String(row[M.コード]||''),sku:M.SKU>=0?String(row[M.SKU]||''):'',qty});
  }
  const candidates=取り置き_初期候補_(orders,取り置き_受注番号集合_(HIKIATE_CFG.部分),取り置き_受注番号集合_(HIKIATE_CFG.希望));
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates);
  ui.alert('取り置き初期登録を作成しました','候補'+candidates.length+'行です。棚の現物を確認し「現物取り置き数量」だけ入力してください。',ui.ButtonSet.OK);
}

function 取り置き初期登録を確定(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  const plan=取り置き_初期確定計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('初期登録を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const selected=inputs.filter(r=>(Number(r.現物取り置き数量)||0)>0), qty=selected.reduce((s,r)=>s+(Number(r.現物取り置き数量)||0),0);
  const answer=ui.alert('初期取り置きを確定します','対象'+selected.length+'行 / 合計'+qty+'個です。確定しますか？',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows);
  SpreadsheetApp.getActive().toast('初期取り置き '+selected.length+'行 / '+qty+'個を確定しました','取り置き台帳',7);
}
