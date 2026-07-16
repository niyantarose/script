const TORIOKI_STATUS = Object.freeze({
  ACTIVE:'取り置き中', SHIPPED:'発送済み', RETURN:'キャンセル戻し', RELEASED:'手動解除'
});
const TORIOKI_RETURN = Object.freeze({
  UNCHECKED:'未確認', PRESENT:'現物あり', REALLOCATED:'再引当済み', YAHOO:'Yahoo反映済み', MISSING:'在庫なし'
});

function 取り置き_整数_(value){
  const n=Number(value);
  return Number.isInteger(n) && n>0 ? n : 0;
}

function 取り置き_商品コード_(sku, code){
  const c=normCode_(code);
  if(c) return c;
  return normCode_(sku).replace(/[AB]$/,'');
}

function 取り置き_行キー_(order){
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

function 取り置き_集計_(rows, movements){
  const out={activeByKey:{}, activeRowsByKey:{}, usageBySupply:{}, usageBySupplyOwner:{}, confirmedReturns:[], errors:[], rows:(rows||[]).map(r=>Object.assign({},r))};
  const ids=new Set(), moveIds=new Set();
  const usage=k=>out.usageBySupply[k]||(out.usageBySupply[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
  const ownerUsage=k=>out.usageBySupplyOwner[k]||(out.usageBySupplyOwner[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
  out.rows.forEach((r,index)=>{
    const id=String(r.取置ID||'').trim(), qty=取り置き_整数_(r.取り置き数量), key=取り置き_行キー_(r);
    if(!id) out.errors.push('台帳'+(index+2)+'行: 取置IDなし');
    else if(ids.has(id)) out.errors.push('台帳'+(index+2)+'行: 取置ID重複 '+id);
    else ids.add(id);
    if(!qty) out.errors.push('台帳'+(index+2)+'行: 取り置き数量は正の整数');
    if(r.状態===TORIOKI_STATUS.ACTIVE && qty){
      out.activeByKey[key]=(out.activeByKey[key]||0)+qty;
      (out.activeRowsByKey[key]=out.activeRowsByKey[key]||[]).push(r);
    }
    const ems=String(r.元EMS番号||'').trim();
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

function 取り置き_決定ID_(source, ems, sourceCode, key, originId){
  const sourceIdentity=取り置き_供給キー_('',sourceCode).slice(1);
  return [source,String(ems||''),sourceIdentity,String(originId||''),key].join('|');
}

function 取り置き_新規行_(order, qty, source, ems, originId, sourceCode){
  const key=取り置き_行キー_(order);
  return {
    取置ID:取り置き_決定ID_(source,ems,sourceCode||order.code,key,originId), 状態:TORIOKI_STATUS.ACTIVE,
    受注番号:String(order.ban), 商品コード:取り置き_商品コード_(order.sku,order.code), SKU:String(order.sku||''),
    取り置き数量:qty, 取置元種別:source, 元EMS番号:String(ems||''), 元EMS商品コード:String(sourceCode||order.code||'').trim(), 元取置ID:String(originId||''),
    戻し処理結果:'', 終了理由・メモ:''
  };
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
  if(/^\d{7,}$/.test(normCode_(code)) && owner.length===1) return owner[0];
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
      : {取置ID:source.取置ID,戻し処理結果:TORIOKI_RETURN.PRESENT,取り置き数量:left,終了理由・メモ:(originalQty-left)+'個を再引当済み'});
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
  supplies.filter(s=>s.directBan).forEach(s=>{
    const available=remainingBySupply[s._key]||0;
    if(available<=0) return;
    const order=取り置き_direct注文_(orders,s.directBan,s.code);
    if(!order) errors.push('注文番号指定の対象なし: '+s.directBan);
    else if(takeSupply(s,order,available,'EMS')!==available) errors.push('注文番号指定が注文数または供給数を超過: '+s.directBan);
  });
  (input.explicit||[]).forEach(e=>{
    const order=orders.find(o=>String(o.ban)===String(e.ban) && matches(o,e.code));
    const sourceCode=String(e.sourceCode||'').trim();
    const candidates=supplies.filter(s=>!s.directBan && String(s.ems)===String(e.ems) && normCode_(s.code)===normCode_(e.code) &&
      (!sourceCode || 取り置き_供給キー_(s.ems,s.sourceCode)===取り置き_供給キー_(e.ems,sourceCode)));
    const supply=candidates.length===1?candidates[0]:null;
    if(!order || !supply) errors.push('P列確定を特定できない: '+e.ban+' '+e.code);
    else {
      const requested=Number(e.qty)||0, taken=takeSupply(supply,order,requested,'EMS');
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
  const plan={orders,newRows,returnUpdates,remainingBySupply,surplus,errors};
  plan.errors=Array.from(new Set(plan.errors.concat(取り置き_割当検証_(plan,input,summary))));
  return plan;
}

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
  const projectedSummary=取り置き_集計_(projected,input.movements||[]), errors=projectedSummary.errors.slice();
  const orderQty={};
  (input.orders||[]).forEach(order=>{
    const key=取り置き_行キー_(order); orderQty[key]=(orderQty[key]||0)+(Number(order.qty)||0);
  });
  Object.keys(projectedSummary.activeByKey).forEach(key=>{
    if(!(key in orderQty)) errors.push('取り置き中の対象注文なし: '+key);
    else if(projectedSummary.activeByKey[key]>orderQty[key]) errors.push('注文数量超過: '+key+' 取り置き'+projectedSummary.activeByKey[key]+' / 注文'+orderQty[key]);
  });
  const supplyQty={};
  (input.supplies||[]).forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.sourceCode||s.code); supplyQty[key]=(supplyQty[key]||0)+(Number(s.qty)||0);
  });
  Object.keys(projectedSummary.usageBySupply).forEach(key=>{
    if(!(key in supplyQty)) return; // 締め済み過去EMSは全件検算で確認し、現在④の供給上限には使わない
    const used=取り置き_使用合計_(projectedSummary.usageBySupply[key]), supplied=supplyQty[key]||0;
    if(used>supplied) errors.push('EMS供給超過: '+key+' 使用'+used+' / 供給'+supplied);
  });
  (plan.newRows||[]).forEach(row=>{
    if(!取り置き_整数_(row.取り置き数量)) errors.push('新規取り置き数量が正の整数でない: '+row.取置ID);
    const order=(input.orders||[]).find(o=>取り置き_行キー_(o)===取り置き_行キー_(row));
    if(order && 取り置き_商品コード_(order.sku,order.code)!==normCode_(row.商品コード)) errors.push('数値枝番または商品コード不一致: '+row.取置ID);
  });
  return Array.from(new Set(errors));
}
