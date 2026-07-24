// ===== 商品診断: 商品コード単位で「箱・注文・台帳・履歴」を全部並べて1個単位で検算する =====
// 「30個着たのに数えたら29のはず」のような食い違いを、どの消費者が食っているかまで見えるようにする。
// 結果は「商品診断」シートに書く(読み取り専用・何も変更しない)。

function 商品診断(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const resp=ui.prompt('商品診断','調べたい商品コードを入力（例 MRBLUE40-7）',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const q=String(resp.getResponseText()||'').trim();
  if(!q) return;
  const targetKeys=[];
  受注候補コード_(q,q).forEach(v=> codeKeys_(v).forEach(k=>{ if(targetKeys.indexOf(k)<0) targetKeys.push(k); }));
  const 一致_=(sku,code)=>{ const keys=[];
    受注候補コード_(sku,code).forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); }));
    return keys.some(k=>targetKeys.indexOf(k)>=0); };

  const out=[]; // [ラベル, A, B, C, D]
  const push=(a,b,c,d,e)=>out.push([a,b,c,d,e]);

  // ── EMSリストの箱(発注共有・全状態) ──
  let 供給=0; const 到着日Set={};
  push('■ EMSリストの箱','状態','到着日','数量','P列(名指し)');
  try{
    const ems=個別対応_EMSリスト_();
    if(ems.error){ push('(EMSリストが読めません)','','','',''); }
    else ems.rows.forEach(r=>{
      if(!一致_('', r.商品コード)) return;
      push('箱', r.状態, r.EMS到着日, r.数量, r.注文番号||'(空)');
      if(String(r.状態||'').trim()==='到着済'){ 供給+=Number(r.数量)||0; const d=ymd_(r.EMS到着日); if(d) 到着日Set[d]=true; }
    });
  }catch(e){ push('(EMSリスト確認でエラー)','','','',''); }

  // ── 受注明細 ──
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  let スタンプ一致=0, 未着=0;
  push('','','','','');
  push('■ 受注明細','受注番号','入荷日','個数','判定');
  if(recv){
    const M=列マップ_(recv), R=recv.getDataRange().getValues();
    for(let i=M.hr;i<R.length;i++){
      const row=R[i];
      const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
      const code=String(row[M.コード]||'').trim(), sku=M.SKU>=0?String(row[M.SKU]||'').trim():'';
      if(!一致_(sku,code)) continue;
      if(区分_(row[M.選択肢])!=='取り寄せ') continue;
      const qty=Number(row[M.個数])||0; if(qty<=0) continue;
      const 入荷v=M.入荷>=0? row[M.入荷] : '';
      const d=String(入荷v||'').trim()? ymd_(入荷v) : '';
      let 判定;
      if(!d){ 判定='未着(待ち)'; 未着+=qty; }
      else if(到着日Set[d]){ 判定='この便で消費'; スタンプ一致+=qty; }
      else 判定='別便(今回対象外)';
      if(!入金済み_(row[M.入金])) 判定+='・未入金';
      push('注文(行'+(i+1)+')', ban, d||'', qty, 判定);
    }
  }

  // ── 消込台帳(表に出ない消費者) ──
  let 出荷済一致=0;
  push('','','','','');
  push('■ 消込台帳','受注番号','入荷日','個数','判定');
  try{
    const tsh=ss.getSheetByName(KESHIKOMI_CFG.シート);
    if(tsh && tsh.getLastRow()>1){
      const tv=tsh.getRange(2,1,tsh.getLastRow()-1,KESHIKOMI_CFG.HDR.length).getDisplayValues();
      tv.forEach(r=>{
        const code=String(r[1]||'').trim(), sku=String(r[2]||'').trim();
        if(!一致_(sku,code)) return;
        const 状態=String(r[5]||'').trim(), qty=Number(r[3])||0, d=ymd_(r[4]);
        let 判定='対象外';
        if(状態.indexOf('出荷済み')===0 && d && 到着日Set[d]){ 判定='★今回の箱を消費(表に出ない)'; 出荷済一致+=qty; }
        push('台帳', String(r[0]||''), d||String(r[4]||''), qty, 状態+' / '+判定);
      });
    }
  }catch(e){ push('(台帳確認でエラー)','','','',''); }

  // ── 引当履歴 ──
  push('','','','','');
  push('■ 引当履歴','受注番号','取込区分/EMS状態','引当数','状態');
  try{
    const hsh=ss.getSheetByName(HIKIATE_HISTORY_CFG.シート);
    if(hsh && hsh.getLastRow()>1){
      const hc=引当履歴_列_(hsh);
      const hv=hsh.getRange(2,1,hsh.getLastRow()-1,hsh.getLastColumn()).getDisplayValues();
      hv.forEach(r=>{
        const code=hc['商品コード']?String(r[hc['商品コード']-1]||''):'';
        if(!一致_('',code)) return;
        const kind=hc['取込区分']?r[hc['取込区分']-1]:'', st=hc['EMSリスト状態']?r[hc['EMSリスト状態']-1]:'',
              ban=hc['受注番号']?r[hc['受注番号']-1]:'', qq=hc['引当数']?r[hc['引当数']-1]:'',
              state=hc['状態']?String(r[hc['状態']-1]||''):'';
        push('履歴', ban, kind+'/'+st, qq, (state||'有効')+'(記録のみ)');
      });
    }
  }catch(e){ push('(履歴確認でエラー)','','','',''); }

  // ── 集計 ──
  push('','','','','');
  push('■ 集計(到着済ベース)','','','','');
  push('到着済の供給','','',供給,'');
  push('この便スタンプ(受注明細)','','',スタンプ一致,'※過去箱分は履歴の差引きで一部相殺されます');
  push('出荷済み消費(台帳)','','',出荷済一致,'今回入荷EMSの在庫には表示されない消費者');
  push('未着の待ち','','',未着,'');
  push('見込み余り(供給-スタンプ-出荷済)','','',供給-スタンプ一致-出荷済一致,'マイナス=どこかで二重、想定と違えば上の一覧で犯人を特定');

  // ── シートへ ──
  const NAME='商品診断';
  let rep=ss.getSheetByName(NAME); if(!rep) rep=ss.insertSheet(NAME);
  rep.clearContents();
  rep.getRange(1,1).setValue('商品診断: '+q+'（照合キー: '+targetKeys.join(', ')+'） '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm'));
  rep.getRange(2,1,out.length,5).setValues(out).setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(1);
  ss.setActiveSheet(rep);
  ss.toast('商品診断: '+q+' → 供給'+供給+' / スタンプ'+スタンプ一致+' / 出荷済'+出荷済一致+' / 未着'+未着,'🔎商品診断',8);
}
