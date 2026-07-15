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
  return String(ems||'').trim()+'|'+normCode_(code);
}

function 取り置き_集計_(rows, movements){
  const out={activeByKey:{}, activeRowsByKey:{}, usageBySupply:{}, confirmedReturns:[], errors:[], rows:(rows||[]).map(r=>Object.assign({},r))};
  const ids=new Set(), moveIds=new Set();
  const usage=k=>out.usageBySupply[k]||(out.usageBySupply[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
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
    const u=usage(取り置き_供給キー_(ems,r.元EMS商品コード||r.商品コード));
    if(r.状態===TORIOKI_STATUS.ACTIVE) u.取り置き中+=qty;
    else if(r.状態===TORIOKI_STATUS.SHIPPED) u.発送済み+=qty;
    else if(r.状態===TORIOKI_STATUS.RETURN){
      const result=String(r.戻し処理結果||TORIOKI_RETURN.UNCHECKED);
      if(result===TORIOKI_RETURN.UNCHECKED || result===TORIOKI_RETURN.PRESENT) u.戻し未処理+=qty;
      if(result===TORIOKI_RETURN.MISSING) u.在庫なし確定+=qty;
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
