const TORIOKI_STATUS = Object.freeze({
  ACTIVE:'取り置き中', SHIPPED:'発送済み', RETURN:'キャンセル戻し', RELEASED:'手動解除'
});
const TORIOKI_RETURN = Object.freeze({
  UNCHECKED:'未確認', PRESENT:'現物あり', REALLOCATED:'再引当済み', YAHOO:'Yahoo反映済み', MISSING:'在庫なし'
});
const TORIOKI_STAGE = Object.freeze({PLANNED:'先行',ARRIVED:'到着済',PHYSICAL:'現物確認済み'});

function 取り置き_整数_(value){
  const n=Number(value);
  return Number.isInteger(n) && n>0 ? n : 0;
}

function 取り置き_商品コード_(sku, code){
  const c=normCode_(code);
  if(c) return c;
  return normCode_(sku).replace(/[AB]$/,'');
}

// 行キー = 受注番号|照合コード。照合コードはSKU優先(枝番a/bもそのまま)、SKUが無い商品だけ商品コード。
// 商品コードとSKUの両方を比較に使うと親コード/バリエーションの表記差で照合が外れるため使わない(2026-07-21)
function 取り置き_行キー_(order){
  const ban=String(order&&order.ban||order&&order.受注番号||'').trim();
  const sku=normCode_(order&&order.sku||order&&order.SKU);
  const code=normCode_(order&&order.code||order&&order.商品コード);
  return ban+'|'+(sku||code);
}

// 取置IDの組成部(従来の3部形式)。台帳に保存済みのID(INIT|/即納|/別ルート|)との互換のため、
// 照合用の行キーとは分離して従来形式を維持する(IDは不透明な識別子・照合には使わない)
function 取り置き_ID部_(order){
  return [
    String(order&&order.ban||order&&order.受注番号||'').trim(),
    取り置き_商品コード_(order&&order.sku||order&&order.SKU, order&&order.code||order&&order.商品コード),
    normCode_(order&&order.sku||order&&order.SKU)
  ].join('|');
}

function 取り置き_供給キー_(ems, code){
  const sourceCode=String(code==null?'':code).trim().toUpperCase().replace(/_/g,'-');
  return String(ems||'').trim()+'|'+sourceCode;
}

function 取り置き_EMS状態_(emsStatusByNo, ems){
  const value=emsStatusByNo instanceof Map ? emsStatusByNo.get(String(ems||'')) : (emsStatusByNo||{})[String(ems||'')];
  const text=String(value&&value.状態||value&&value.status||value||'').trim();
  // 未着の証拠がある箱だけ先行。空/不明/在庫反映済みの旧箱は従来どおり物理確保(到着済)へ倒す
  // (状態不明を先行に落とすと、旧台帳のEMS行が一斉に「物理に無い」扱いになり出荷が止まる)
  if(/未着/.test(text)) return TORIOKI_STAGE.PLANNED;
  if(text==='') return TORIOKI_STAGE.ARRIVED;
  return /到着|入荷|現物|反映済|arrived/i.test(text) ? TORIOKI_STAGE.ARRIVED : TORIOKI_STAGE.PLANNED;
}
function 取り置き_段階正規化_(row, emsStatusByNo){
  const r=row||{}, explicit=String(r.引当段階||'').trim();
  if([TORIOKI_STAGE.PLANNED,TORIOKI_STAGE.ARRIVED,TORIOKI_STAGE.PHYSICAL].indexOf(explicit)>=0) return Object.assign({},r,{引当段階:explicit});
  if(explicit==='要移行') return Object.assign({},r,{引当段階:'要移行'});
  if(String(r.取置元種別||'')==='開始前在庫') return Object.assign({},r,{引当段階:'要移行'});
  if(String(r.元EMS番号||'').trim()) return Object.assign({},r,{引当段階:取り置き_EMS状態_(emsStatusByNo,r.元EMS番号)});
  // 旧台帳互換: 段階列も元EMS番号も無いが種別のある取り置き中(EMS/キャンセル再引当/手動等)は、
  // 従来どおり現在の確保として数える(=到着済)。種別まで無い行だけ要確認へ。
  if(String(r.取置元種別||'').trim()) return Object.assign({},r,{引当段階:TORIOKI_STAGE.ARRIVED});
  return Object.assign({},r,{引当段階:'要確認'});
}
function 取り置き_段階別集計_(rows, movements, emsStatusByNo){
  const out={byKey:{},activeByKey:{},activeRowsByKey:{},usageBySupply:{},要移行行:[],要確認行:[]};
  (rows||[]).forEach(row=>{
    if(row.状態!==TORIOKI_STATUS.ACTIVE) return;
    const r=取り置き_段階正規化_(row,emsStatusByNo),qty=取り置き_整数_(r.取り置き数量),key=取り置き_行キー_(r);
    if(!qty) return;
    if(r.引当段階==='要移行'){
      // 旧開始前在庫は移行完了まで「確保」として数える(activeByKeyから消すと④の残必要が二重割当する)。
      // 段階バケットへは混ぜず、要移行数量として分離して持つ(移行UI・台帳確保集計の除外用)。
      out.要移行行.push(r);
      const g=out.byKey[key]||(out.byKey[key]={現物確認済み数量:0,到着済引当数量:0,先行引当数量:0,合計確保数量:0,要移行数量:0,行内訳:[]});
      g.要移行数量=(g.要移行数量||0)+qty;
      out.activeByKey[key]=(out.activeByKey[key]||0)+qty;
      (out.activeRowsByKey[key]=out.activeRowsByKey[key]||[]).push(r);
      return;
    }
    if(r.引当段階==='要確認'){out.要確認行.push(r);return;}
    const group=out.byKey[key]||(out.byKey[key]={現物確認済み数量:0,到着済引当数量:0,先行引当数量:0,合計確保数量:0,要移行数量:0,行内訳:[]});
    if(r.引当段階===TORIOKI_STAGE.PHYSICAL) group.現物確認済み数量+=qty;
    else if(r.引当段階===TORIOKI_STAGE.ARRIVED) group.到着済引当数量+=qty; else group.先行引当数量+=qty;
    group.合計確保数量+=qty; group.行内訳.push({取置ID:String(r.取置ID||''),引当段階:r.引当段階,数量:qty});
    out.activeByKey[key]=(out.activeByKey[key]||0)+qty;(out.activeRowsByKey[key]=out.activeRowsByKey[key]||[]).push(r);
    const ems=取り置き_実効供給EMS_(r,r.引当段階); if(ems){const supplyKey=取り置き_供給キー_(ems,r.元EMS商品コード||r.商品コード);out.usageBySupply[supplyKey]=(out.usageBySupply[supplyKey]||0)+qty;}
  }); return out;
}

function 取り置き_集計_(rows, movements, emsStatusByNo){
  const stages=取り置き_段階別集計_(rows,movements,emsStatusByNo);
  const out={activeByKey:stages.activeByKey,activeRowsByKey:stages.activeRowsByKey,stageByKey:stages.byKey,要移行行:stages.要移行行,要確認行:stages.要確認行,usageBySupply:{},usageBySupplyOwner:{},confirmedReturns:[],errors:[],rows:(rows||[]).map(r=>Object.assign({},r))};
  const ids=new Set(), moveIds=new Set();
  const usage=k=>out.usageBySupply[k]||(out.usageBySupply[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
  const ownerUsage=k=>out.usageBySupplyOwner[k]||(out.usageBySupplyOwner[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
  out.rows.forEach((r,index)=>{
    const id=String(r.取置ID||'').trim(), qty=取り置き_整数_(r.取り置き数量), key=取り置き_行キー_(r);
    if(!id) out.errors.push('台帳'+(index+2)+'行: 取置IDなし');
    else if(ids.has(id)) out.errors.push('台帳'+(index+2)+'行: 取置ID重複 '+id);
    else ids.add(id);
    if(!qty) out.errors.push('台帳'+(index+2)+'行: 取り置き数量は正の整数');
    const stage=取り置き_段階正規化_(r,emsStatusByNo);
    const ems=r.状態===TORIOKI_STATUS.ACTIVE ? 取り置き_実効供給EMS_(r,stage.引当段階) : String(r.元EMS番号||'').trim();
    if(!ems || !qty || r.状態===TORIOKI_STATUS.RELEASED) return;
    const supplyKey=取り置き_供給キー_(ems,r.元EMS商品コード||r.商品コード);
    const u=usage(supplyKey),owned=ownerUsage(supplyKey+'|'+String(r.受注番号||'').trim());
    if(r.状態===TORIOKI_STATUS.ACTIVE){ u.取り置き中+=qty; owned.取り置き中+=qty; }
    else if(r.状態===TORIOKI_STATUS.SHIPPED){ u.発送済み+=qty; owned.発送済み+=qty; }
    else if(r.状態===TORIOKI_STATUS.RETURN){
      const result=String(r.戻し処理結果||TORIOKI_RETURN.UNCHECKED);
      if(result===TORIOKI_RETURN.UNCHECKED || result===TORIOKI_RETURN.PRESENT){ u.戻し未処理+=qty; owned.戻し未処理+=qty; }
      if(result===TORIOKI_RETURN.MISSING){ u.在庫なし確定+=qty; owned.在庫なし確定+=qty; }
      if(result===TORIOKI_RETURN.PRESENT) out.confirmedReturns.push(r);
    }
  });
  (movements||[]).forEach((m,index)=>{
    const qty=取り置き_整数_(m.数量), id=String(m.処理ID||'').trim();
    if(!id || !qty) out.errors.push('移動台帳'+(index+2)+'行: 処理IDまたは数量が不正');
    else if(moveIds.has(id)) out.errors.push('移動台帳'+(index+2)+'行: 処理ID重複 '+id);
    else moveIds.add(id);
    if(qty) usage(取り置き_供給キー_(m.EMS番号,m.商品コード)).Yahoo移動済み+=qty;
  });
  return out;
}

function 取り置き_今回必要数_(order, summary){
  const qty=Math.max(0,Number(order&&order.qty)||0);
  return Math.max(0,qty-Number(summary&&summary.activeByKey[取り置き_行キー_(order)]||0));
}

// 台帳がEMS引当・キャンセル再引当で既に確保している数量(行キー単位)。
// 開始前在庫は取り置き登録シート自身の確定分(洗い替えで数量を引き継いで再確定する)なので含めない。
// 発送済み・手動解除・キャンセル戻しは現在の確保ではないため数えない(取り置き中のみ)。
// 取り置き登録画面の「④確保分」。開始前在庫(要移行)は登録シート自身の確定分なので従来どおり除く。
function 取り置き_台帳確保集計_(ledgerRows){
  const stages=取り置き_段階別集計_(ledgerRows||[],[],{});
  const out={};
  Object.keys(stages.activeByKey).forEach(key=>{
    const 要移行=(stages.byKey[key]&&stages.byKey[key].要移行数量)||0;
    const qty=(stages.activeByKey[key]||0)-要移行;
    if(qty>0) out[key]=qty;
  });
  return out;
}

function 取り置き_決定ID_(source, ems, sourceCode, key, originId){
  const sourceIdentity=取り置き_供給キー_('',sourceCode).slice(1);
  return [source,String(ems||''),sourceIdentity,String(originId||''),key].join('|');
}

function 取り置き_新規行_(order, qty, source, ems, originId, sourceCode){
  const key=取り置き_行キー_(order);
  const id=取り置き_決定ID_(source,ems,sourceCode||order.code,key,originId);
  return {
    取置ID:id, 状態:TORIOKI_STATUS.ACTIVE,
    受注番号:String(order.ban), 商品コード:取り置き_商品コード_(order.sku,order.code), SKU:String(order.sku||''),
    // ④のEMS引当・キャンセル再引当は到着済箱の現物からだけ作られるため段階は到着済(先行は再計算エンジンだけが作る)
    取り置き数量:qty, 取置元種別:source, 引当段階:TORIOKI_STAGE.ARRIVED, 引当系譜ID:id, 引当系譜数量:qty, 元EMS番号:String(ems||''), 元EMS商品コード:String(sourceCode||order.code||'').trim(), 元取置ID:String(originId||''),
    // 「・」はGASのpush時パーサーが識別子として受け付けない(Nodeは通る)ため、キーは必ず引用符で囲む
    戻し処理結果:'', '終了理由・メモ':''
  };
}

function 取り置き_供給状態マップ_(supplies,now){
  const out={},time=now instanceof Date?now.getTime():Date.now();
  // 到着予定日は予定であり実到着の証拠にしない。明示状態(未着/到着済)と到着確認フラグを優先し、
  // 日付で推定してよいのは実到着日(arrival/到着日)だけ。
  (supplies||[]).forEach(s=>{const ems=String(s.ems||s.EMS番号||'').trim(),st=String(s.状態||s.status||'');
    const 明示到着=s.isArrived===true||/到着|入荷|現物|arrived/i.test(st),明示未着=/未着/.test(st);
    const arrival=Date.parse(String(s.arrival||s.到着日||''));
    out[ems]=明示到着||(!明示未着&&isFinite(arrival)&&arrival<=time)?'到着済':'未着';});return out;
}
function 取り置き_実効供給EMS_(row,stage){return stage===TORIOKI_STAGE.PHYSICAL&&row.供給処理==='供給解放'?'':(stage===TORIOKI_STAGE.PHYSICAL?String(row.供給控除EMS||row.元EMS番号||'').trim():String(row.元EMS番号||'').trim());}
function 取り置き_系譜ID_(row){return String(row.引当系譜ID||row.取置ID||'').trim();}
function 取り置き_系譜数量_(row){return 取り置き_整数_(row.引当系譜数量)||取り置き_整数_(row.取り置き数量);}
function 取り置き_不変条件検証_(orders,ledgerRows,supplies){
  const errors=[],warnings=[],status=取り置き_供給状態マップ_(supplies,new Date()),ordersByKey={},supplyQty={},supplyUsed={},lineage={},ids=new Set();
  (orders||[]).forEach(o=>{const key=取り置き_行キー_(o);ordersByKey[key]=(ordersByKey[key]||0)+(Number(o.qty||o.個数)||0);});
  (supplies||[]).forEach(s=>{const key=取り置き_供給キー_(s.ems||s.EMS番号,s.sourceCode||s.code||s.商品コード);supplyQty[key]=(supplyQty[key]||0)+(Number(s.qty||s.数量)||0);});
  const stages=取り置き_段階別集計_(ledgerRows||[],[],status);
  (ledgerRows||[]).forEach((row,index)=>{
    const id=String(row.取置ID||'').trim(),qty=取り置き_整数_(row.取り置き数量),stage=取り置き_段階正規化_(row,status).引当段階;
    if(!id) errors.push('台帳'+(index+2)+'行: 取置IDなし');else if(ids.has(id)) errors.push('取置ID重複: '+id);else ids.add(id);
    if(row.状態!==TORIOKI_STATUS.ACTIVE||!qty||stage==='要移行'||stage==='要確認') return;
    const lineageId=取り置き_系譜ID_(row),lineageQty=取り置き_系譜数量_(row),l=lineage[lineageId]||(lineage[lineageId]={qty:0,limit:lineageQty});l.qty+=qty;l.limit=Math.max(l.limit,lineageQty);
    const ems=取り置き_実効供給EMS_(row,stage);if(ems){const key=取り置き_供給キー_(ems,row.元EMS商品コード||row.商品コード);supplyUsed[key]=(supplyUsed[key]||0)+qty;}
  });
  Object.keys(lineage).forEach(id=>{if(lineage[id].qty>lineage[id].limit)errors.push('段階重複: '+id+' 有効'+lineage[id].qty+' / 系譜'+lineage[id].limit);});
  Object.keys(stages.activeByKey).forEach(key=>{const reserved=stages.activeByKey[key],physical=stages.byKey[key].現物確認済み数量;if(!(key in ordersByKey)){if(physical)errors.push('現物固定の注文が消滅: '+key);else warnings.push('取り置き台帳に現注文の無い行: '+key);}else{if(reserved>ordersByKey[key])errors.push('注文数を超過: '+key);if(physical>ordersByKey[key])errors.push('現物固定の注文数量減: '+key);}});
  Object.keys(supplyUsed).forEach(key=>{if(!(key in supplyQty))errors.push('EMS供給不在: '+key.replace('|',' '));else if(supplyUsed[key]>supplyQty[key])errors.push('EMS供給超過: '+key.replace('|',' '));});
  stages.要移行行.forEach(r=>warnings.push('旧開始前在庫は要移行: '+String(r.取置ID||'')));return {errors:Array.from(new Set(errors)),warnings:Array.from(new Set(warnings))};
}
// 供給処理を明示し、元EMS不明現物は到着済供給を複数箱FIFOで分割する。
function 取り置き_現物確認変換計画_(inputs,ledger,supplies,now){
  const original=(ledger||[]).map(r=>Object.assign({},r)),at=now||new Date(),status=取り置き_供給状態マップ_(supplies,at),targets={},errors=[],review=[];
  // 空欄・空白のみは「未入力」= その注文SKUに触らない(0=解除要求とは区別する)
  (inputs||[]).forEach(i=>{const raw=i.現物取り置き数量;if(raw==null||String(raw).trim()==='')return;const k=取り置き_行キー_(i),q=取り置き_整数_(raw);if(targets[k])errors.push('現物確認の重複入力: '+k);else targets[k]={q,i};});
  Object.keys(targets).forEach(k=>{const a=original.filter(r=>r.状態===TORIOKI_STATUS.ACTIVE&&取り置き_行キー_(r)===k),p=a.filter(r=>取り置き_段階正規化_(r,status).引当段階===TORIOKI_STAGE.PHYSICAL).reduce((n,r)=>n+取り置き_整数_(r.取り置き数量),0),t=a.reduce((n,r)=>n+取り置き_整数_(r.取り置き数量),0);if(targets[k].q<p)errors.push('解除フローを使用してください: '+k);if(targets[k].q>t)errors.push('新規加算はできません: '+k);targets[k].convert=targets[k].q-p;});if(errors.length)return{rows:original,errors,review};
  const cap={},used=取り置き_段階別集計_(original,[],status).usageBySupply;(supplies||[]).forEach(s=>{const k=取り置き_供給キー_(s.ems,s.code);cap[k]=(cap[k]||0)+(Number(s.qty)||0);});Object.keys(used).forEach(k=>cap[k]=Math.max(0,(cap[k]||0)-used[k]));
  const ids=new Set(original.map(r=>String(r.取置ID||''))),next=(id)=>{let n=1,v;do{v=id+'|現物|'+n++;}while(ids.has(v));ids.add(v);return v;},out=[];
  for(const r of original){const t=targets[取り置き_行キー_(r)],stage=取り置き_段階正規化_(r,status),q=取り置き_整数_(r.取り置き数量);if(!t||stage.引当段階===TORIOKI_STAGE.PHYSICAL||r.状態!==TORIOKI_STATUS.ACTIVE||!t.convert){out.push(r);continue;}const take=Math.min(q,t.convert);t.convert-=take;let parts=[];
    if(r.元EMS番号)parts=[{ems:r.元EMS番号,qty:take,process:stage.引当段階===TORIOKI_STAGE.ARRIVED?'EMS控除':'供給解放'}];else{let left=take;for(const s of (supplies||[]).slice().sort((a,b)=>String(a.arrival).localeCompare(String(b.arrival))||String(a.ems).localeCompare(String(b.ems)))){const k=取り置き_供給キー_(s.ems,s.code),n=Math.min(left,cap[k]||0);if(取り置き_EMS状態_(status,s.ems)===TORIOKI_STAGE.ARRIVED&&取り置き_商品コード_('',s.code)===取り置き_商品コード_(r.SKU,r.商品コード)&&n){parts.push({ems:s.ems,qty:n,process:'EMS控除'});cap[k]-=n;left-=n;}}if(left){review.push({理由:'到着済供給控除EMSを特定できません'});return{rows:original,errors,review};}}
    if(q>take)out.push(Object.assign({},r,{取り置き数量:q-take}));parts.forEach((p,i)=>out.push(Object.assign({},r,{取置ID:q===take&&i===0?r.取置ID:next(r.取置ID),取り置き数量:p.qty,引当段階:TORIOKI_STAGE.PHYSICAL,供給控除EMS:p.process==='EMS控除'?p.ems:'',供給処理:p.process,引当系譜ID:取り置き_系譜ID_(r),引当系譜数量:取り置き_系譜数量_(r),現物確認日時:at,現物確認メモ:String(t.i.現物確認メモ||r.現物確認メモ||'')})));
  }return{rows:out,errors,review};
}
function 取り置き_使用合計_(usage){
  return ['取り置き中','発送済み','戻し未処理','在庫なし確定','Yahoo移動済み']
    .reduce((sum,key)=>sum+(Number(usage&&usage[key])||0),0);
}

function 取り置き_注文照合_(order,code){
  const matchCode=normCode_(code),keys=order&&order.keys instanceof Set?Array.from(order.keys):(order&&order.keys||[]);
  return keys.indexOf(matchCode)>=0 || 取り置き_商品コード_(order&&order.sku,order&&order.code)===matchCode;
}

function 取り置き_direct注文_(orders,directBan,code){
  const owner=(orders||[]).filter(o=>String(o.ban||'')===String(directBan||''));
  const matched=owner.filter(o=>取り置き_注文照合_(o,code));
  if(matched.length===1) return matched[0];
  // コードがどの行とも一致しない名指し(受注番号コード・説明文コード+P列手動名指し等)は、
  // その注文の取り寄せ行が1行だけなら救済する。複数行は曖昧なので呼び出し側が警告して余りへ。
  if(matched.length===0 && owner.length===1) return owner[0];
  return null;
}

function P列計画_純計算_(emsRows, inputOrders, fixedBySupply, usageBySupply){
  const orders=(inputOrders||[]).map(o=>Object.assign({},o));
  const rows=(emsRows||[]).map(r=>Object.assign({},r,{entries:[],left:0,nextP:''}));
  const fixedRemaining={}, usageRemaining={};
  Object.keys(fixedBySupply||{}).forEach(key=>{
    fixedRemaining[key]=(fixedBySupply[key]||[]).map(e=>Object.assign({},e));
  });
  Object.keys(usageBySupply||{}).forEach(key=>{
    usageRemaining[key]=取り置き_使用合計_(usageBySupply[key]);
  });
  const takeFixed=(key,limit,directBan)=>{
    const out=[], queue=fixedRemaining[key]||[]; let left=limit,index=0;
    while(left>0 && index<queue.length){
      if(directBan && String(queue[index].ban)!==String(directBan)){ index++; continue; }
      const take=Math.min(left,Number(queue[index].qty)||0);
      if(take>0) out.push({ban:String(queue[index].ban),qty:take,explicit:false,direct:!!directBan});
      queue[index].qty-=take; left-=take;
      if(queue[index].qty<=0) queue.splice(index,1); else index++;
    }
    return out;
  };
  const fifo=orders.slice().sort((a,b)=>a.date-b.date||(a.seq||0)-(b.seq||0));
  rows.forEach(r=>{
    r.sourceCode=String(r.sourceCode||r.code||'').trim();
    r.directBan=String(r.directBan||'').trim();
    const key=取り置き_供給キー_(r.ems,r.sourceCode), qty=Math.max(0,Number(r.qty)||0);
    const fixed=takeFixed(key,qty,r.directBan);
    const fixedQty=fixed.reduce((sum,e)=>sum+(Number(e.qty)||0),0);
    usageRemaining[key]=Math.max(0,(usageRemaining[key]||0)-fixedQty);
    const nonDisplayUsed=Math.min(qty-fixedQty,usageRemaining[key]||0);
    usageRemaining[key]=Math.max(0,(usageRemaining[key]||0)-nonDisplayUsed);
    let capacity=Math.max(0,qty-fixedQty-nonDisplayUsed);
    r.entries=fixed;
    if(r.directBan){
      const order=取り置き_direct注文_(fifo,r.directBan,r.code);
      if(order && capacity>0 && order.need>0){
        const take=Math.min(capacity,order.need); capacity-=take; order.need-=take;
        const prev=r.entries.find(e=>String(e.ban)===r.directBan);
        if(prev) prev.qty+=take;
        else r.entries.push({ban:r.directBan,qty:take,explicit:false,direct:true});
      }
      r.left=capacity;
      r.nextP=P列指定文字列_(r.entries,qty);
      return;
    }
    for(const order of fifo){
      if(capacity<=0) break;
      const keys=order.keys instanceof Set?Array.from(order.keys):(order.keys||[]);
      if(keys.indexOf(normCode_(r.code))<0 || !(order.need>0)) continue;
      const take=Math.min(capacity,order.need); capacity-=take; order.need-=take;
      const prev=r.entries.find(e=>String(e.ban)===String(order.ban));
      if(prev) prev.qty+=take;
      else r.entries.push({ban:String(order.ban),qty:take,explicit:false});
    }
    r.left=capacity;
    r.nextP=P列指定文字列_(r.entries,qty);
  });
  return {rows,orders};
}

function 取り置き_割当計算_(input){
  const summary=取り置き_集計_(input.ledger||[],input.movements||[]);
  const orders=(input.orders||[]).map(o=>Object.assign({},o,{need:取り置き_今回必要数_(o,summary)}))
    .sort((a,b)=>(a.sortKey||0)-(b.sortKey||0)||(a.i||0)-(b.i||0));
  const newRows=[], newRowById=Object.create(null), returnUpdates=[], errors=summary.errors.slice();
  const addNewRow=row=>{
    const id=String(row.取置ID||''), existing=newRowById[id];
    if(existing) existing.取り置き数量+=row.取り置き数量;
    else { newRowById[id]=row; newRows.push(row); }
  };
  const matches=(order,code)=>取り置き_注文照合_(order,code);
  summary.confirmedReturns.slice().sort((a,b)=>String(a.登録日時||'').localeCompare(String(b.登録日時||''))).forEach(source=>{
    const originalQty=取り置き_整数_(source.取り置き数量); let left=originalQty;
    for(const order of orders){
      if(!left || !order.need || !matches(order,source.商品コード)) continue;
      const take=Math.min(left,order.need); order.need-=take; left-=take;
      addNewRow(取り置き_新規行_(order,take,'キャンセル再引当',source.元EMS番号,source.取置ID,source.元EMS商品コード||source.商品コード));
    }
    if(left<originalQty) returnUpdates.push(left===0
      ? {取置ID:source.取置ID,戻し処理結果:TORIOKI_RETURN.REALLOCATED}
      : {取置ID:source.取置ID,戻し処理結果:TORIOKI_RETURN.PRESENT,取り置き数量:left,'終了理由・メモ':(originalQty-left)+'個を再引当済み'});
  });
  const supplyByKey={};
  (input.supplies||[]).forEach(s=>{
    const sourceCode=String(s.sourceCode||s.code||'').trim(),directBan=String(s.directBan||'').trim();
    const sourceKey=取り置き_供給キー_(s.ems,sourceCode),key=sourceKey+'|'+directBan;
    if(!supplyByKey[key]) supplyByKey[key]=Object.assign({},s,{code:normCode_(s.code),sourceCode,directBan,qty:0,_key:key,_sourceKey:sourceKey});
    supplyByKey[key].qty+=(Number(s.qty)||0);
  });
  const supplies=Object.keys(supplyByKey).map(key=>supplyByKey[key]).sort((a,b)=>
    String(a.arrival||'').localeCompare(String(b.arrival||'')) ||
    String(a.ems||'').localeCompare(String(b.ems||'')) ||
    String(a.sourceCode||'').localeCompare(String(b.sourceCode||'')) ||
    String(a.directBan||'').localeCompare(String(b.directBan||'')));
  const remainingBySupply={};
  const unassignedUsage={};
  supplies.forEach(s=>{
    if(!(s._sourceKey in unassignedUsage)) unassignedUsage[s._sourceKey]=取り置き_使用合計_(summary.usageBySupply[s._sourceKey]);
    const owned=s.directBan?取り置き_使用合計_(summary.usageBySupplyOwner[s._sourceKey+'|'+s.directBan]):0;
    const used=Math.min(Number(s.qty)||0,owned,unassignedUsage[s._sourceKey]||0);
    remainingBySupply[s._key]=Math.max(0,(Number(s.qty)||0)-used);
    unassignedUsage[s._sourceKey]=Math.max(0,(unassignedUsage[s._sourceKey]||0)-used);
  });
  supplies.forEach(s=>{
    const used=Math.min(remainingBySupply[s._key]||0,unassignedUsage[s._sourceKey]||0);
    remainingBySupply[s._key]=Math.max(0,(remainingBySupply[s._key]||0)-used);
    unassignedUsage[s._sourceKey]=Math.max(0,(unassignedUsage[s._sourceKey]||0)-used);
  });
  const takeSupply=(s,order,qty,source)=>{
    const key=s._key, take=Math.min(qty,order.need,remainingBySupply[key]||0);
    if(take<=0) return 0;
    remainingBySupply[key]-=take; order.need-=take;
    addNewRow(取り置き_新規行_(order,take,source,s.ems,'',s.sourceCode));
    return take;
  };
  // タグの名指し先が不在(キャンセル済み等)・数量超過(多め買付)は日常運用で起きるため、
  // エラー(④全体中止)ではなく警告にして持ち主のない分を余り(日本在庫)へ回す。
  const warnings=[];
  supplies.filter(s=>s.directBan).forEach(s=>{
    const available=remainingBySupply[s._key]||0;
    if(available<=0) return;
    const label=String(s.sourceCode||s.code||'').trim()||normCode_(s.code);
    const order=取り置き_direct注文_(orders,s.directBan,s.code);
    if(!order) warnings.push('注文番号指定の対象なし(全量を余りへ): '+s.directBan+' '+label+' x'+available+' → キャンセル済みなら発注共有EMSリストのタグを外してください');
    else {
      const taken=takeSupply(s,order,available,'EMS');
      if(taken!==available) warnings.push('注文番号指定の余剰を余りへ: '+s.directBan+' '+label+' 供給'+available+'／引当'+taken);
    }
  });
  (input.explicit||[]).forEach(e=>{
    // 同一商品を複数行に分けた注文(分割行)があるため、最初の1行ではなく一致する全行へ順に配る
    // (実例 2026-07-21: 10117725 MEDAFB01W 2個+1個の2行に1個×3箱の名指し→3個目が割当0で④中止)
    const matched=orders.filter(o=>String(o.ban)===String(e.ban) && matches(o,e.code));
    const sourceCode=String(e.sourceCode||'').trim();
    const candidates=supplies.filter(s=>!s.directBan && String(s.ems)===String(e.ems) && normCode_(s.code)===normCode_(e.code) &&
      (!sourceCode || 取り置き_供給キー_(s.ems,s.sourceCode)===取り置き_供給キー_(e.ems,sourceCode)));
    const supply=candidates.length===1?candidates[0]:null;
    if(!matched.length || !supply) errors.push('P列確定を特定できない: '+e.ban+' '+e.code);
    else {
      const requested=Number(e.qty)||0;
      let taken=0;
      for(const order of matched){
        if(taken>=requested) break;
        taken+=takeSupply(supply,order,requested-taken,'EMS');
      }
      if(taken!==requested) errors.push('P列確定数量を満たせない: '+e.ban+' '+e.code+' 指定'+requested+' / 割当'+taken);
    }
  });
  supplies.filter(s=>!s.directBan).forEach(s=>{
    for(const order of orders){
      if(!(remainingBySupply[s._key]>0)) break;
      if(!order.need) continue;
      if(matches(order,s.code)) takeSupply(s,order,order.need,'EMS');
    }
  });
  const surplus=supplies.map(s=>({ems:s.ems,code:s.sourceCode,matchCode:normCode_(s.code),sourceCode:s.sourceCode,directBan:s.directBan,
    qty:remainingBySupply[s._key]||0,arrival:s.arrival}))
    .filter(s=>s.qty>0);
  const plan={orders,newRows,returnUpdates,remainingBySupply,surplus,errors,warnings:Array.from(new Set(warnings))};
  const 検証=取り置き_割当検証_(plan,input);
  plan.errors=Array.from(new Set(plan.errors.concat(検証.errors)));
  plan.warnings=Array.from(new Set(plan.warnings.concat(検証.warnings)));
  return plan;
}

// ===== 個別引当・引当キャンセルボタン用(台帳直書き) =====
// P列は④が計画から書き直す派生表示になったため、ボタンは台帳へ直接書く。

// 到着済みEMS行から、台帳使用を差し引いた「残あり」候補を返す。
// targetKeys: 受注側の照合キー(受注候補コード_+codeKeys_で広めに)。ban一致の受注番号タグは
// コード不一致でも候補(REQ系・名指し買付の救済はこのボタンが担う)。
function 個別_台帳候補_(emsRows, targetKeys, ban, summary){
  const keys=(targetKeys||[]).map(k=>normCode_(k)).filter(Boolean);
  const out=[];
  (emsRows||[]).forEach(r=>{
    if(String(r.状態||'').trim()!=='到着済') return;
    const ems=String(r.EMS番号||'').trim();
    if(!実EMS番号_(ems)) return;
    const tagHit=!!(ban && タグ受注番号_(r.商品コード)===String(ban));
    if(!tagHit && !codeKeys_(r.商品コード).some(k=>keys.indexOf(k)>=0)) return;
    const supplyKey=取り置き_供給キー_(ems, r.商品コード);
    const 残=Math.max(0,(Number(r.数量)||0)-取り置き_使用合計_(summary&&summary.usageBySupply?summary.usageBySupply[supplyKey]:null));
    if(残>0) out.push({item:r, 残});
  });
  return out;
}

// 台帳へ取り置き中行を積む(同じ決定ID=同じ注文×箱なら数量を加算)。保存は呼び出し側。
function 個別_台帳引当行_(ledgerRows, order, take, ems, sourceCode, now){
  const row=取り置き_新規行_(order, take, 'EMS', ems, '', sourceCode);
  const rows=(ledgerRows||[]).map(r=>Object.assign({},r));
  const index=rows.findIndex(r=>String(r.取置ID||'')===row.取置ID);
  if(index>=0){
    const cur=rows[index];
    if(cur.状態!==TORIOKI_STATUS.ACTIVE) return {rows:null,error:'同じ取置IDの行が取り置き中ではありません: '+row.取置ID};
    rows[index]=Object.assign({},cur,{取り置き数量:取り置き_整数_(cur.取り置き数量)+take,更新日時:now});
  } else {
    rows.push(Object.assign({},row,{登録日時:now,更新日時:now}));
  }
  return {rows,error:''};
}

// 行キーの取り置き中だけを手動解除へ。発送済み・他注文は触らない。保存は呼び出し側。
function 個別_台帳解除計画_(ledgerRows, key, reason, now){
  let released=0, qty=0;
  const rows=(ledgerRows||[]).map(r=>{
    if(r.状態!==TORIOKI_STATUS.ACTIVE || 取り置き_行キー_(r)!==key) return Object.assign({},r);
    released++; qty+=取り置き_整数_(r.取り置き数量);
    return Object.assign({},r,{状態:TORIOKI_STATUS.RELEASED,'終了理由・メモ':reason,更新日時:now});
  });
  return {rows,released,qty};
}

// 台帳に載らない出荷(週末GoQ直接発送など④未実行の箱からの出荷)を検知し、
// 「発送済み」行として自動登録する計画を作る。旧9a8f8d8の売り越し防止の台帳版。
//   shippedRows: 消込台帳_出荷済み行_()の行 / ledgerRows・movements: 現台帳
//   supplies: EMS供給オブジェクト_の行(到着済の実EMSのみ)
// ルール: 台帳に行キーが既にある注文・メモ「取り置き」・発送日より後に着いた箱は対象外。
//         既存使用を差し引いた箱残の範囲だけ登録し、届かない分は要確認へ返す。
function 取り置き_未台帳出荷計画_(shippedRows, ledgerRows, movements, supplies){
  const summary=取り置き_集計_(ledgerRows||[], movements||[]);
  const ledgerKeys=new Set((ledgerRows||[]).map(r=>取り置き_行キー_(r)));
  const targets={}, order=[];
  (shippedRows||[]).forEach(r=>{
    if(取り置き出荷_(r)) return; // メモ「取り置き」= 開始前在庫から出荷。箱を食わせない
    const qty=Math.max(0,Number(r.qty)||0); if(!qty) return;
    const key=取り置き_行キー_({ban:r.ban,code:r.code,sku:r.sku});
    if(ledgerKeys.has(key)) return; // 台帳が知っている注文はCSV遷移側で扱う
    if(!targets[key]){ targets[key]={ban:String(r.ban||'').trim(),code:String(r.code||''),sku:String(r.sku||''),qty:0,ship:''}; order.push(key); }
    targets[key].qty+=qty;
    const d=ymd_(r.基準日); if(d && (!targets[key].ship || d<targets[key].ship)) targets[key].ship=d;
  });
  const groups={}; // sourceKey -> {ems,sourceCode,matchCode,directBan,arrival(最古),qty}
  (supplies||[]).forEach(s=>{
    const sourceCode=String(s.sourceCode||s.code||'').trim(), key=取り置き_供給キー_(s.ems,sourceCode);
    if(!groups[key]) groups[key]={_key:key,ems:String(s.ems||''),sourceCode,matchCode:normCode_(s.code),
      directBan:String(s.directBan||'').trim(),arrival:String(s.arrival||''),qty:0};
    groups[key].qty+=(Number(s.qty)||0);
    const a=String(s.arrival||''); if(a && (!groups[key].arrival || a<groups[key].arrival)) groups[key].arrival=a;
  });
  const remaining={};
  Object.keys(groups).forEach(key=>{
    remaining[key]=Math.max(0,groups[key].qty-取り置き_使用合計_(summary.usageBySupply[key]));
  });
  const sorted=Object.keys(groups).map(k=>groups[k]).sort((a,b)=>
    String(a.arrival).localeCompare(String(b.arrival)) || String(a.ems).localeCompare(String(b.ems)));
  const newRows=[], review=[];
  order.forEach(key=>{
    const t=targets[key]; let left=t.qty;
    const keys=引当用照合キー一覧_(t.sku,t.code);
    for(const g of sorted){
      if(left<=0) break;
      if(g.directBan && g.directBan!==t.ban) continue;           // 他注文の名指し箱は食わない
      // 商品一致、またはこの注文への名指し箱(説明文コード等はコード不一致でも出どころ)
      if(keys.indexOf(g.matchCode)<0 && g.directBan!==t.ban) continue;
      if(t.ship && g.arrival && g.arrival>t.ship) continue;      // 発送より後に着いた箱は出どころではない
      const take=Math.min(left, remaining[g._key]||0); if(take<=0) continue;
      remaining[g._key]-=take; left-=take;
      const row=取り置き_新規行_({ban:t.ban,code:t.code,sku:t.sku}, take, '出荷実績', g.ems, '', g.sourceCode);
      row.状態=TORIOKI_STATUS.SHIPPED;
      row['終了理由・メモ']='④自動登録(台帳外出荷)';
      newRows.push(row);
    }
    if(left>0 && left<t.qty) review.push({受注番号:t.ban,商品コード:取り置き_商品コード_(t.sku,t.code),
      理由:'台帳外出荷のうち'+(t.qty-left)+'/'+t.qty+'だけ箱残と一致(残りは供給不足)'});
  });
  return {newRows,review};
}

// 戻り {errors, warnings}:
//   errors  = ④を止めるべき異常(台帳の重複ID・非整数、計画が作った不正行=内部バグ)
//   warnings= ④は完走し人に知らせる注意(注文が消えた/減った台帳行、箱の供給超過)。
//     注文が消えた・減った取り置きは、現物は既に確保済みなので勝手に消さず「手動解除で調整」を促す。
//     供給超過は⑤便締めをブロックする突合超過(引当.js)と同種なので、ここでも警告に留める。
function 取り置き_割当検証_(plan,input){
  const projected=(input.ledger||[]).map(r=>Object.assign({},r));
  const byId={}; projected.forEach((r,index)=>byId[String(r.取置ID||'')]=index);
  (plan.returnUpdates||[]).forEach(update=>{
    const index=byId[String(update.取置ID||'')];
    if(index!==undefined) projected[index]=Object.assign({},projected[index],update);
  });
  (plan.newRows||[]).forEach(row=>{
    const id=String(row.取置ID||''), index=byId[id];
    if(index===undefined){ byId[id]=projected.length; projected.push(Object.assign({},row)); }
    else projected[index]=Object.assign({},projected[index],row);
  });
  const projectedSummary=取り置き_集計_(projected,input.movements||[]);
  const errors=projectedSummary.errors.slice(), warnings=[];
  projectedSummary.要移行行.forEach(row=>warnings.push('旧開始前在庫は要移行(通常引当から除外): '+取り置き_行キー_(row)));
  const orderQty={};
  (input.orders||[]).forEach(order=>{
    const key=取り置き_行キー_(order); orderQty[key]=(orderQty[key]||0)+(Number(order.qty)||0);
  });
  Object.keys(projectedSummary.activeByKey).forEach(key=>{
    const reserved=projectedSummary.activeByKey[key];
    if(!(key in orderQty)) warnings.push('取り置き台帳に現注文の無い行(出荷/削除済みなら手動解除してください): '+key+' 取り置き'+reserved);
    else if(reserved>orderQty[key]) warnings.push('取り置きが注文数を超過(手動解除で調整してください): '+key+' 取り置き'+reserved+' / 注文'+orderQty[key]);
  });
  const supplyQty={};
  (input.supplies||[]).forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.sourceCode||s.code); supplyQty[key]=(supplyQty[key]||0)+(Number(s.qty)||0);
  });
  Object.keys(projectedSummary.usageBySupply).forEach(key=>{
    if(!(key in supplyQty)) return; // 締め済み過去EMSは全件検算で確認し、現在④の供給上限には使わない
    const used=取り置き_使用合計_(projectedSummary.usageBySupply[key]), supplied=supplyQty[key]||0;
    if(used>supplied) warnings.push('EMS供給超過(便締め前に要確認): '+key.replace('|',' ')+' 使用'+used+' / 供給'+supplied);
  });
  (plan.newRows||[]).forEach(row=>{
    if(!取り置き_整数_(row.取り置き数量)) errors.push('新規取り置き数量が正の整数でない: '+row.取置ID);
    const order=(input.orders||[]).find(o=>取り置き_行キー_(o)===取り置き_行キー_(row));
    if(order && 取り置き_商品コード_(order.sku,order.code)!==normCode_(row.商品コード)) errors.push('数値枝番または商品コード不一致: '+row.取置ID);
  });
  return {errors:Array.from(new Set(errors)), warnings:Array.from(new Set(warnings))};
}
