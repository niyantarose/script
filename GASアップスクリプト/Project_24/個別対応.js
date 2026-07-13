// ===== 個別対応: 照合キーからEMSリストを補完し、個別引当/キャンセルを整合的に実行 =====

const KOBETSU_HIKIATE_CFG = {
  シート: '個別対応',
  HDR: ['操作','照合キー','受注番号','商品コード','数量','引当数','EMS番号','Box No.','購入No.','EMS到着日','現在の注文番号','理由','実行結果','実行日時'],
  ボタン行: 1,
  引当ボタン列: 16,
  取消ボタン列: 18,
  ボタン幅: 150,
  ボタン高さ: 36
};

function 個別対応シートを作成(){
  個別対応_シート_();
  SpreadsheetApp.getActive().toast('個別対応シートを作成/確認しました','個別対応',5);
}

function 個別受注引当(){
  個別対応_実行_('個別引当');
}

function 引当キャンセル(){
  個別対応_実行_('引当キャンセル');
}

function 個別対応_照合キーから補完(){
  const r=個別対応_補完_(false);
  if(r.error){ SpreadsheetApp.getUi().alert(r.error); return; }
  SpreadsheetApp.getActive().toast('個別対応 補完: '+r.補完+'行 / エラー '+r.エラー+'行','個別対応',6);
}

function 個別対応_シート_(){
  const ss=SpreadsheetApp.getActive(), cfg=KOBETSU_HIKIATE_CFG;
  let sh=ss.getSheetByName(cfg.シート), 新規=!sh;
  if(!sh) sh=ss.insertSheet(cfg.シート);
  sh.getRange(1,1,1,cfg.HDR.length).setValues([cfg.HDR])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#4472c4')
    .setHorizontalAlignment('center')
    .setFontSize(HIKIATE_CFG.字);
  sh.getRange(1,5).setNote('EMSリストのその行の数量です。照合キーから自動で入ります。');
  sh.getRange(1,6).setNote('この受注番号へ割り当てる数量です。空欄ならEMS行の残り全部を割り当てます。');
  sh.setFrozenRows(1);
  if(新規){
    [100,260,110,210,70,70,150,70,150,110,180,200,260,150].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  }
  const rule=SpreadsheetApp.newDataValidation().requireValueInList(['個別引当','引当キャンセル'], true).setAllowInvalid(true).build();
  sh.getRange(2,1,Math.max(1,sh.getMaxRows()-1),1).setDataValidation(rule);
  個別対応ボタンを設置_(true);
  return sh;
}

function 個別対応_列_(sh){
  const head=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),KOBETSU_HIKIATE_CFG.HDR.length)).getDisplayValues()[0].map(v=>String(v||'').trim());
  const out={}; KOBETSU_HIKIATE_CFG.HDR.forEach(h=>out[h]=head.indexOf(h)+1);
  return out;
}

function 個別対応ボタンを設置(){
  個別対応_シート_();
  SpreadsheetApp.getActive().toast('個別対応ボタンを設置しました','個別対応',5);
}

function 個別対応ボタンを設置_(silent){
  const sh=SpreadsheetApp.getActive().getSheetByName(KOBETSU_HIKIATE_CFG.シート);
  if(!sh){ if(!silent) SpreadsheetApp.getUi().alert('個別対応シートがありません。'); return 0; }
  const cfg=KOBETSU_HIKIATE_CFG;
  個別対応_旧チェックボタンをクリア_(sh);
  個別対応_既存画像ボタン削除_(sh);
  try{
    個別対応_画像ボタン追加_(sh,cfg.引当ボタン列,'個別受注引当','#6aa84f','#38761d','#ffffff','個別受注引当');
    個別対応_画像ボタン追加_(sh,cfg.取消ボタン列,'引当キャンセル','#cc0000','#990000','#ffffff','引当キャンセル');
  }catch(e){
    個別対応_セルボタンを設置_(sh);
  }
  sh.setRowHeight(cfg.ボタン行, Math.max(sh.getRowHeight(cfg.ボタン行), cfg.ボタン高さ+8));
  if(sh.getColumnWidth(cfg.引当ボタン列)<cfg.ボタン幅) sh.setColumnWidth(cfg.引当ボタン列,cfg.ボタン幅);
  if(sh.getColumnWidth(cfg.取消ボタン列)<cfg.ボタン幅) sh.setColumnWidth(cfg.取消ボタン列,cfg.ボタン幅);
  return 1;
}

function 個別対応_旧チェックボタンをクリア_(sh){
  const cfg=KOBETSU_HIKIATE_CFG;
  [cfg.引当ボタン列,cfg.引当ボタン列+1,cfg.取消ボタン列,cfg.取消ボタン列+1].forEach(col=>{
    sh.getRange(cfg.ボタン行,col).clearDataValidations().clearContent().setBackground(null).setNote('');
  });
}

function 個別対応_既存画像ボタン削除_(sh){
  try{
    sh.getImages().forEach(img=>{
      const title=String(img.getAltTextTitle&&img.getAltTextTitle()||'');
      if(title.indexOf('個別対応ボタン_')===0) img.remove();
    });
  }catch(e){}
}

function 個別対応_ボタン画像Blob_(label, fill, stroke, textColor){
  const cfg=KOBETSU_HIKIATE_CFG;
  const esc=s=>String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const svg='<svg xmlns="http://www.w3.org/2000/svg" width="'+cfg.ボタン幅+'" height="'+cfg.ボタン高さ+'" viewBox="0 0 '+cfg.ボタン幅+' '+cfg.ボタン高さ+'">'
    +'<rect x="1" y="1" width="'+(cfg.ボタン幅-2)+'" height="'+(cfg.ボタン高さ-2)+'" rx="5" ry="5" fill="'+fill+'" stroke="'+stroke+'" stroke-width="2"/>'
    +'<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Arial, sans-serif" font-size="15" font-weight="700" fill="'+textColor+'">'+esc(label)+'</text>'
    +'</svg>';
  return Utilities.newBlob(svg, 'image/svg+xml', '個別対応ボタン_'+label+'.svg');
}

function 個別対応_画像ボタン追加_(sh, col, label, fill, stroke, textColor, fn){
  const cfg=KOBETSU_HIKIATE_CFG;
  const img=sh.insertImage(個別対応_ボタン画像Blob_(label,fill,stroke,textColor), col, cfg.ボタン行);
  img.setAltTextTitle('個別対応ボタン_'+label);
  img.setAltTextDescription(label+'を実行');
  img.assignScript(fn);
  img.setWidth(cfg.ボタン幅);
  img.setHeight(cfg.ボタン高さ);
}

function 個別対応_セルボタンを設置_(sh){
  const cfg=KOBETSU_HIKIATE_CFG;
  sh.getRange(cfg.ボタン行,cfg.引当ボタン列).setValue('個別受注引当')
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#6aa84f')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('クリックしてセルを選択すると個別受注引当を実行します。');
  sh.getRange(cfg.ボタン行,cfg.取消ボタン列).setValue('引当キャンセル')
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#cc0000')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('クリックしてセルを選択すると引当キャンセルを実行します。');
}

function 個別対応セルボタン選択_(e){
  if(!e || !e.range) return false;
  const sh=e.range.getSheet(), cfg=KOBETSU_HIKIATE_CFG;
  if(sh.getName()!==cfg.シート || e.range.getRow()!==cfg.ボタン行) return false;
  const text=String(e.range.getDisplayValue()||'').trim();
  if(e.range.getColumn()===cfg.引当ボタン列 && text==='個別受注引当'){
    個別受注引当();
    return true;
  }
  if(e.range.getColumn()===cfg.取消ボタン列 && text==='引当キャンセル'){
    引当キャンセル();
    return true;
  }
  return false;
}

function 個別対応_数値_(v){
  const n=Number(String(v==null?'':v).replace(/[０-９．]/g,d=>String.fromCharCode(d.charCodeAt(0)-0xFEE0)).replace(/[^\d.\-]/g,''));
  return isNaN(n)?0:n;
}

function 個別対応_P注文展開_(text, rowQty){
  const out=[], fallback=rowQty>0?rowQty:1;
  String(text||'').split(/[,、]/).forEach(part=>{
    const m=String(part||'').trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
    if(m) out.push({ban:m[1], qty:m[2]?Number(m[2]):fallback});
  });
  return out;
}

function 個別対応_P注文整形_(entries, rowQty){
  const order=[], map={};
  entries.forEach(e=>{
    const ban=String(e.ban||'').trim(), qty=個別対応_数値_(e.qty);
    if(!ban || qty<=0) return;
    if(!(ban in map)){ map[ban]=0; order.push(ban); }
    map[ban]+=qty;
  });
  const rows=order.map(ban=>({ban, qty:map[ban]})).filter(e=>e.qty>0);
  if(!rows.length) return '';
  if(rows.length===1 && rowQty>0 && rows[0].qty===rowQty) return rows[0].ban;
  return rows.map(e=>e.ban+':'+e.qty).join(', ');
}

function 個別対応_EMSリスト_(){
  const cfg=P_KAKUTEI_CFG;
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!sh) return {error:'発注共有ファイルに「'+cfg.シート+'」がありません'};
  const hr=cfg.ヘッダー行, last=sh.getLastRow();
  if(last<=hr) return {error:'EMSリストにデータがありません'};
  const lastCol=sh.getLastColumn();
  const head=sh.getRange(hr,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const c=引当履歴_EMSリスト列_(head);
  if(c.照合キー<0 || c.注文番号<0) return {error:'EMSリストに「照合キー」または「注文番号」見出しがありません'};
  const data=sh.getRange(hr+1,1,last-hr,lastCol).getDisplayValues();
  const rows=data.map((r,i)=>({
    row: hr+1+i,
    状態: String(r[c.状態]||'').trim(),
    EMS到着日: String(r[c.到着日]||'').trim(),
    購入No: String(r[c.購入No]||'').trim(),
    商品コード: String(r[c.商品コード]||'').trim(),
    数量: 個別対応_数値_(r[c.数量]),
    EMS番号: String(r[c.EMS番号]||'').trim(),
    BoxNo: String(r[c.BoxNo]||'').trim(),
    照合キー: String(r[c.照合キー]||'').trim(),
    注文番号: String(r[c.注文番号]||'').trim()
  })).filter(r=>実EMS番号_(r.EMS番号));
  return {sh, c, rows};
}

function 個別対応_EMS行を探す_(ems, req, mode){
  let rows=ems.rows.filter(r=>r.照合キー===req.照合キー);
  if(req.EMS番号) rows=rows.filter(r=>r.EMS番号===req.EMS番号);
  if(req.BoxNo) rows=rows.filter(r=>r.BoxNo===req.BoxNo);
  if(mode==='個別引当'){
    const arrived=rows.filter(r=>r.状態==='到着済');
    if(arrived.length) rows=arrived;
    else return {error:'照合キーに一致する「到着済」のEMSリスト行がありません'};
  }else if(req.受注番号){
    const withBan=rows.filter(r=>個別対応_P注文展開_(r.注文番号,r.数量).some(e=>e.ban===req.受注番号));
    if(withBan.length) rows=withBan;
  }
  if(!rows.length) return {error:'照合キーに一致するEMSリスト行がありません'};
  if(rows.length>1) return {error:'候補が複数あります。EMS番号かBox No.も入れてください'};
  return {item:rows[0]};
}

function 個別対応_履歴Rec_(item, ban, qty, kind){
  return {
    取込区分: kind,
    EMSリスト状態: item.状態,
    EMSリスト行: item.row,
    EMS番号: item.EMS番号,
    BoxNo: item.BoxNo,
    EMS到着日: item.EMS到着日,
    購入No: item.購入No,
    照合キー: item.照合キー,
    商品コード: item.商品コード,
    EMS行数量: item.数量,
    受注番号: ban,
    引当数: qty
  };
}

function 個別対応_補完_(quiet){
  const sh=個別対応_シート_(), c=個別対応_列_(sh), last=sh.getLastRow();
  if(last<=1) return {補完:0, エラー:0};
  const ems=個別対応_EMSリスト_();
  if(ems.error) return {error:ems.error};
  const vals=sh.getRange(2,1,last-1,KOBETSU_HIKIATE_CFG.HDR.length).getDisplayValues();
  let 補完=0, エラー=0;
  vals.forEach((r,i)=>{
    const rowNo=2+i, key=String(r[c['照合キー']-1]||'').trim();
    if(!key) return;
    const mode=String(r[c['操作']-1]||'').trim()||'個別引当';
    const req={照合キー:key, 受注番号:String(r[c['受注番号']-1]||'').trim(), EMS番号:String(r[c['EMS番号']-1]||'').trim(), BoxNo:String(r[c['Box No.']-1]||'').trim()};
    const found=個別対応_EMS行を探す_(ems, req, mode);
    if(found.error){
      エラー++;
      if(!quiet && c['実行結果']) sh.getRange(rowNo,c['実行結果']).setValue(found.error);
      return;
    }
    個別対応_行補完_(sh,rowNo,c,found.item);
    補完++;
  });
  return {補完, エラー};
}

function 個別対応_行補完_(sh, rowNo, c, item){
  const set=(name,val)=>{ if(c[name]) sh.getRange(rowNo,c[name]).setValue(val); };
  set('商品コード', item.商品コード);
  set('数量', item.数量);
  set('EMS番号', item.EMS番号);
  set('Box No.', item.BoxNo);
  set('購入No.', item.購入No);
  set('EMS到着日', item.EMS到着日);
  set('現在の注文番号', item.注文番号);
}

function 個別対応_実行_(mode){
  const ss=SpreadsheetApp.getActive(), sh=個別対応_シート_(), c=個別対応_列_(sh), last=sh.getLastRow();
  if(last<=1){ ss.toast('個別対応: 実行する行がありません','個別対応',5); return; }
  const ems=個別対応_EMSリスト_();
  if(ems.error){ SpreadsheetApp.getUi().alert(ems.error); return; }
  const vals=sh.getRange(2,1,last-1,KOBETSU_HIKIATE_CFG.HDR.length).getDisplayValues();
  let ok=0, ng=0, skip=0;
  vals.forEach((r,i)=>{
    const rowNo=2+i;
    const result=String(r[c['実行結果']-1]||'').trim();
    if(result){ skip++; return; }
    const op=String(r[c['操作']-1]||'').trim();
    if(op && op!==mode){ skip++; return; }
    const req={
      照合キー:String(r[c['照合キー']-1]||'').trim(),
      受注番号:String(r[c['受注番号']-1]||'').trim(),
      EMS番号:String(r[c['EMS番号']-1]||'').trim(),
      BoxNo:String(r[c['Box No.']-1]||'').trim()
    };
    if(!req.照合キー && !req.受注番号){ skip++; return; }
    const done=(msg,isOk)=>{
      sh.getRange(rowNo,c['実行結果']).setValue(msg);
      sh.getRange(rowNo,c['実行日時']).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      if(isOk) ok++; else ng++;
    };
    if(!req.照合キー){ done('照合キーがありません',false); return; }
    if(!req.受注番号){ done('受注番号がありません',false); return; }
    const found=個別対応_EMS行を探す_(ems, req, mode);
    if(found.error){ done(found.error,false); return; }
    const item=found.item;
    個別対応_行補完_(sh,rowNo,c,item);
    if(c['操作'] && !op) sh.getRange(rowNo,c['操作']).setValue(mode);
    const cell=ems.sh.getRange(item.row,ems.c.注文番号+1);
    const entries=個別対応_P注文展開_(item.注文番号,item.数量);
    if(mode==='個別引当'){
      const used=entries.reduce((s,e)=>s+e.qty,0);
      const inputQty=個別対応_数値_(r[c['引当数']-1]);
      const qty=inputQty>0?inputQty:Math.max(0,item.数量-used);
      if(qty<=0){ done('引当できる残数がありません',false); return; }
      if(used+qty>item.数量){ done('EMS行数量を超えます（残 '+Math.max(0,item.数量-used)+'）',false); return; }
      entries.push({ban:req.受注番号, qty});
      const text=個別対応_P注文整形_(entries,item.数量);
      cell.setValue(text);
      cell.setBackground(/[,:：、]/.test(text)?'#fff2cc':null);
      引当履歴_個別記録_(個別対応_履歴Rec_(item, req.受注番号, qty, '個別引当'));
      done('OK 個別引当: '+req.受注番号+' x'+qty,true);
      return;
    }
    const current=entries.filter(e=>e.ban===req.受注番号).reduce((s,e)=>s+e.qty,0);
    const inputQty=個別対応_数値_(r[c['引当数']-1]);
    const qty=inputQty>0?inputQty:current;
    if(qty<=0){ done('キャンセル対象の引当が見つかりません',false); return; }
    let left=qty;
    const kept=entries.map(e=>{
      if(e.ban!==req.受注番号 || left<=0) return e;
      const take=Math.min(e.qty,left);
      left-=take;
      return {ban:e.ban, qty:e.qty-take};
    }).filter(e=>e.qty>0);
    const text=個別対応_P注文整形_(kept,item.数量);
    cell.setValue(text);
    cell.setBackground(/[,:：、]/.test(text)?'#fff2cc':null);
    const canceled=qty-left;
    const hist=引当履歴_キャンセル_(個別対応_履歴Rec_(item, req.受注番号, qty, '個別引当'), qty);
    const cleared=個別対応_入荷日クリア_(req.受注番号, item.商品コード, item.EMS到着日);
    if(canceled<=0 && hist.更新<=0){ done('キャンセル対象の引当が見つかりません',false); return; }
    done('OK キャンセル: '+req.受注番号+' x'+(canceled||qty)+' / 履歴'+hist.更新+'件 / 入荷日クリア'+cleared+'行',true);
  });
  ss.toast('個別対応 '+mode+': OK '+ok+' / エラー '+ng+' / スキップ '+skip,'個別対応',8);
}

function 個別対応_同商品有効割当あり_(ban, code){
  const targetKeys=codeKeys_(code);
  const ems=個別対応_EMSリスト_();
  if(!ems.error){
    for(const r of ems.rows){
      if(!r.注文番号) continue;
      const hitCode=codeKeys_(r.商品コード).some(k=>targetKeys.indexOf(k)>=0);
      if(!hitCode) continue;
      if(個別対応_P注文展開_(r.注文番号,r.数量).some(e=>e.ban===ban && e.qty>0)) return true;
    }
  }
  const hist=引当履歴_反映済み割当マップ_();
  return (hist[ban]||[]).some(e=>e.qty>0 && targetKeys.indexOf(e.key)>=0);
}

function 個別対応_日付キー_(v){
  if(v instanceof Date && !isNaN(v.getTime())){
    return v.getFullYear()+'-'+('0'+(v.getMonth()+1)).slice(-2)+'-'+('0'+v.getDate()).slice(-2);
  }
  const s=String(v||'').trim();
  let m=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
  m=s.match(/(\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return '20'+m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
  return '';
}

function 個別対応_入荷日クリア_(ban, code, arrival){
  if(個別対応_同商品有効割当あり_(ban,code)) return 0;
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return 0;
  const M=列マップ_(recv);
  if(M.入荷<0) return 0;
  const R=recv.getDataRange().getValues(), targetKeys=codeKeys_(code), arrivalKey=個別対応_日付キー_(arrival);
  let cleared=0;
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    if(String(row[M.番号]||'').trim()!==ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const hit=codeKeys_(row[M.コード]).some(k=>targetKeys.indexOf(k)>=0);
    if(!hit) continue;
    const cur=row[M.入荷];
    if(!String(cur||'').trim()) continue;
    const curKey=個別対応_日付キー_(cur);
    if(arrivalKey && curKey && arrivalKey!==curKey) continue;
    recv.getRange(i+1,M.入荷+1).clearContent();
    cleared++;
  }
  return cleared;
}
