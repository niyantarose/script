// ===== 引当履歴: EMSリストの「どの箱をどの受注に使ったか」を引当ファイル側にも固定保存 =====
// EMSリストのP列(注文番号)は発注共有ファイル側のメモなので、ここへコピーして二重記録を防ぎつつ後から追えるようにする。

const HIKIATE_HISTORY_CFG = {
  シート: '引当履歴',
  HDR: ['履歴キー','取込区分','EMSリスト状態','EMSリスト行','EMS番号','Box No.','EMS到着日','購入No.','照合キー','商品コード','EMS行数量','受注番号','引当数','記録日時','状態']
};

function 引当履歴シートを作成(){
  引当履歴_シート_();
  SpreadsheetApp.getActive().toast('引当履歴シートを作成/確認しました','🧾引当履歴',5);
}

function 引当履歴_今回到着分を記録(){
  const r=引当履歴_今回到着分を記録_(false);
  if(r.error){ SpreadsheetApp.getUi().alert(r.error); return; }
  SpreadsheetApp.getActive().toast('引当履歴: 今回分 追加'+r.追加+'件 / 重複スキップ'+r.重複+'件 / 対象'+r.対象+'件','🧾引当履歴',7);
}

function 引当履歴_今回到着分を記録_(quiet){
  const r=引当履歴_EMSリストから記録_(['到着済'], '今回到着');
  if(!quiet){
    if(r.error) SpreadsheetApp.getUi().alert(r.error);
    else SpreadsheetApp.getActive().toast('引当履歴: 今回分 追加'+r.追加+'件 / 重複スキップ'+r.重複+'件 / 対象'+r.対象+'件','🧾引当履歴',7);
  }
  return r;
}

function 引当履歴_過去データを取込(){
  const ui=SpreadsheetApp.getUi();
  const ans=ui.alert('引当履歴へ過去データ取込',
    'EMSリストの「在庫反映済み」かつ注文番号ありの行を、過去取込として引当履歴に追加します。\n既に同じ履歴キーがあるものは追加しません。実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  const r=引当履歴_EMSリストから記録_(['在庫反映済み'], '過去取込');
  if(r.error){ ui.alert(r.error); return; }
  SpreadsheetApp.getActive().toast('引当履歴: 過去取込 追加'+r.追加+'件 / 重複スキップ'+r.重複+'件 / 対象'+r.対象+'件','🧾引当履歴',8);
}

function 引当履歴_シート_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_HISTORY_CFG;
  let sh=ss.getSheetByName(cfg.シート), 新規=!sh;
  if(!sh) sh=ss.insertSheet(cfg.シート);
  const lastCol=Math.max(sh.getLastColumn(), cfg.HDR.length);
  const cur=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].slice(0,cfg.HDR.length).join('\t');
  if(cur!==cfg.HDR.join('\t')){
    sh.getRange(1,1,1,cfg.HDR.length).setValues([cfg.HDR])
      .setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setFontSize(HIKIATE_CFG.字);
    sh.setFrozenRows(1);
  }
  if(新規){
    [300,90,110,90,150,70,100,150,260,210,80,110,70,150,100].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  }
  引当履歴_状態列を補完_(sh);
  return sh;
}

function 引当履歴_列_(sh){
  const head=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HIKIATE_HISTORY_CFG.HDR.length)).getDisplayValues()[0].map(v=>String(v||'').trim());
  const out={}; HIKIATE_HISTORY_CFG.HDR.forEach(h=>out[h]=head.indexOf(h)+1);
  return out;
}

function 引当履歴_状態列を補完_(sh){
  const c=引当履歴_列_(sh);
  if(!c['状態'] || sh.getLastRow()<=1) return;
  const n=sh.getLastRow()-1;
  const keys=sh.getRange(2,1,n,1).getDisplayValues();
  const vals=sh.getRange(2,c['状態'],n,1).getDisplayValues();
  let changed=false;
  for(let i=0;i<n;i++){
    if(String(keys[i][0]||'').trim() && !String(vals[i][0]||'').trim()){
      vals[i][0]='有効';
      changed=true;
    }
  }
  if(changed) sh.getRange(2,c['状態'],n,1).setValues(vals);
}

function 引当履歴_既存キー_(sh){
  const set={};
  const last=sh.getLastRow();
  if(last<=1) return set;
  sh.getRange(2,1,last-1,1).getDisplayValues().forEach((r,i)=>{
    const k=String(r[0]||'').trim(); if(k) set[k]=2+i;
  });
  return set;
}

function 引当履歴_EMSリスト列_(head){
  const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  const pick=(idx, fallback)=> idx>=0 ? idx : fallback;
  return {
    状態: pick(f('ステータス列','ステータス'), 6),
    到着日: pick(f('EMS到着日','到着日','到着'), 4),
    購入No: pick(f('購入No.','購入No'), 5),
    商品コード: pick(f('商品コード'), 8),
    数量: pick(f('数量','個数'), 9),
    EMS番号: pick(f('EMS番号'), 12),
    BoxNo: pick(f('Box No.','BoxNo','BOXNo'), 13),
    照合キー: pick(f('照合キー'), 14),
    注文番号: f('注文番号')
  };
}

function 引当履歴_数値_(v){
  const n=Number(String(v==null?'':v).replace(/[^\d.\-]/g,''));
  return isNaN(n)?0:n;
}

function 引当履歴_注文番号展開_(text, rowQty){
  const out=[], fallback=rowQty>0?rowQty:1;
  String(text||'').split(/[,、]/).forEach(part=>{
    const m=String(part||'').trim().match(/^(\d{5,})(?:[:：]\s*(\d+))?$/);
    if(m) out.push({ban:m[1], qty:m[2]?Number(m[2]):fallback});
  });
  return out;
}

function 引当履歴_キー_(rec){
  const shogo=rec.照合キー || (rec.購入No+'_'+normCode_(rec.商品コード));
  return [rec.EMS番号, rec.BoxNo, shogo, rec.購入No, normCode_(rec.商品コード), rec.受注番号]
    .map(v=>String(v==null?'':v).trim()).join('|');
}

function 引当履歴_行_(key, rec, now, state){
  return [key,rec.取込区分,rec.EMSリスト状態,rec.EMSリスト行,rec.EMS番号,rec.BoxNo,rec.EMS到着日,rec.購入No,rec.照合キー,rec.商品コード,rec.EMS行数量,rec.受注番号,rec.引当数,now,state||'有効'];
}

function 引当履歴_EMSリストから記録_(statuses, kind){
  const cfg=P_KAKUTEI_CFG;
  let src;
  try{ src=SpreadsheetApp.openById(cfg.発注共有ID).getSheetByName(cfg.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!src) return {error:'発注共有ファイルに「'+cfg.シート+'」がありません'};

  const hr=cfg.ヘッダー行, last=src.getLastRow();
  if(last<=hr) return {追加:0, 重複:0, 対象:0};
  const lastCol=src.getLastColumn();
  const head=src.getRange(hr,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const c=引当履歴_EMSリスト列_(head);
  if(c.注文番号<0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};

  const data=src.getRange(hr+1,1,last-hr,lastCol).getDisplayValues();
  const targetStatus={}; statuses.forEach(s=>targetStatus[s]=true);
  const sh=引当履歴_シート_(), exists=引当履歴_既存キー_(sh), now=new Date();
  const rows=[], updates=[]; let 重複=0, 対象=0;

  data.forEach((r,i)=>{
    const 状態=String(r[c.状態]||'').trim();
    if(!targetStatus[状態]) return;
    const 注文番号=String(r[c.注文番号]||'').trim();
    if(!注文番号) return;
    const 商品コード=String(r[c.商品コード]||'').trim();
    if(!商品コード) return;
    const rowQty=引当履歴_数値_(r[c.数量]);
    const orders=引当履歴_注文番号展開_(注文番号, rowQty);
    if(!orders.length) return;
    orders.forEach(o=>{
      対象++;
      const rec={
        取込区分: kind,
        EMSリスト状態: 状態,
        EMSリスト行: hr+1+i,
        EMS番号: String(r[c.EMS番号]||'').trim(),
        BoxNo: String(r[c.BoxNo]||'').trim(),
        EMS到着日: String(r[c.到着日]||'').trim(),
        購入No: String(r[c.購入No]||'').trim(),
        照合キー: String(r[c.照合キー]||'').trim(),
        商品コード,
        EMS行数量: rowQty,
        受注番号: o.ban,
        引当数: o.qty
      };
      const key=引当履歴_キー_(rec);
      if(exists[key]){
        重複++;
        if(kind==='過去取込' && exists[key]>=2) updates.push({row:exists[key], kind, status:状態, state:'有効'});
        return;
      }
      exists[key]=-1; // 同じ取込内の重複を防ぐ。既存行ではないので更新対象にはしない。
      rows.push(引当履歴_行_(key,rec,now,'有効'));
    });
  });

  if(rows.length){
    const start=sh.getLastRow()+1;
    sh.getRange(start,1,rows.length,HIKIATE_HISTORY_CFG.HDR.length).setValues(rows)
      .setFontSize(HIKIATE_CFG.字)
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.getRange(start,14,rows.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  const hc=引当履歴_列_(sh);
  updates.forEach(u=>{
    if(hc['取込区分']) sh.getRange(u.row,hc['取込区分']).setValue(u.kind);
    if(hc['EMSリスト状態']) sh.getRange(u.row,hc['EMSリスト状態']).setValue(u.status);
    if(hc['状態']) sh.getRange(u.row,hc['状態']).setValue(u.state||'有効');
  });
  return {追加:rows.length, 重複, 対象};
}

function 引当履歴_個別記録_(rec){
  const sh=引当履歴_シート_(), exists=引当履歴_既存キー_(sh), now=new Date();
  rec.取込区分=rec.取込区分||'個別引当';
  const key=引当履歴_キー_(rec), row=exists[key], c=引当履歴_列_(sh);
  if(row && row>=2){
    const values=引当履歴_行_(key,rec,now,'有効');
    HIKIATE_HISTORY_CFG.HDR.forEach((h,i)=>{
      if(c[h]) sh.getRange(row,c[h]).setValue(values[i]);
    });
    if(c['記録日時']) sh.getRange(row,c['記録日時']).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    return {追加:0, 更新:1, key};
  }
  const start=sh.getLastRow()+1;
  sh.getRange(start,1,1,HIKIATE_HISTORY_CFG.HDR.length).setValues([引当履歴_行_(key,rec,now,'有効')])
    .setFontSize(HIKIATE_CFG.字)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(start,14).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return {追加:1, 更新:0, key};
}

function 引当履歴_キャンセル_(rec, cancelQty){
  const sh=引当履歴_シート_();
  if(sh.getLastRow()<=1) return {更新:0, 残:cancelQty||0};
  const c=引当履歴_列_(sh), lastCol=sh.getLastColumn();
  const vals=sh.getRange(2,1,sh.getLastRow()-1,lastCol).getDisplayValues();
  const key=引当履歴_キー_(rec);
  let left=cancelQty>0?cancelQty:Infinity, updated=0;
  for(let i=0;i<vals.length;i++){
    const row=vals[i], sheetRow=2+i;
    const rowKey=String(row[(c['履歴キー']||1)-1]||'').trim();
    const state=c['状態']?String(row[c['状態']-1]||'').trim():'有効';
    if(state==='キャンセル済み') continue;
    let hit=rowKey===key;
    if(!hit){
      const ban=c['受注番号']?String(row[c['受注番号']-1]||'').trim():'';
      const code=c['商品コード']?normCode_(row[c['商品コード']-1]):'';
      const shogo=c['照合キー']?String(row[c['照合キー']-1]||'').trim():'';
      const ems=c['EMS番号']?String(row[c['EMS番号']-1]||'').trim():'';
      const box=c['Box No.']?String(row[c['Box No.']-1]||'').trim():'';
      hit=ban===String(rec.受注番号||'').trim()
        && (!rec.照合キー || shogo===String(rec.照合キー||'').trim())
        && (!rec.商品コード || code===normCode_(rec.商品コード))
        && (!rec.EMS番号 || ems===String(rec.EMS番号||'').trim())
        && (!rec.BoxNo || box===String(rec.BoxNo||'').trim());
    }
    if(!hit) continue;
    const qty=c['引当数']?引当履歴_数値_(row[c['引当数']-1]):0;
    if(left!==Infinity && qty>left){
      if(c['引当数']) sh.getRange(sheetRow,c['引当数']).setValue(qty-left);
      left=0; updated++;
      break;
    }
    if(c['状態']) sh.getRange(sheetRow,c['状態']).setValue('キャンセル済み');
    left=left===Infinity?Infinity:Math.max(0,left-qty);
    updated++;
    if(left<=0) break;
  }
  return {更新:updated, 残:left===Infinity?0:left};
}

function 引当履歴_反映済み割当マップ_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(HIKIATE_HISTORY_CFG.シート);
  if(!sh || sh.getLastRow()<=1) return {};
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const f=n=>head.indexOf(n);
  const cKind=f('取込区分'), cStatus=f('EMSリスト状態'), cDate=f('EMS到着日'), cCode=f('商品コード'), cBan=f('受注番号'), cQty=f('引当数'), cState=f('状態');
  if(cCode<0 || cBan<0 || cQty<0) return {};
  const vals=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues();
  const map={};
  vals.forEach(r=>{
    const kind=cKind>=0?String(r[cKind]||'').trim():'';
    const status=cStatus>=0?String(r[cStatus]||'').trim():'';
    const state=cState>=0?String(r[cState]||'').trim():'有効';
    if(state==='キャンセル済み') return;
    if(kind!=='過去取込' && status!=='在庫反映済み') return;
    const ban=String(r[cBan]||'').trim(), key=normCode_(r[cCode]), qty=引当履歴_数値_(r[cQty]);
    if(!ban || !key || qty<=0) return;
    (map[ban]=map[ban]||[]).push({key, qty, date:cDate>=0?String(r[cDate]||'').trim():''});
  });
  return map;
}

function 引当履歴_需要を差し引く_(lines){
  const map=引当履歴_反映済み割当マップ_();
  lines.forEach(l=>{
    const q=map[l.ban]||[];
    q.forEach(e=>{
      if(l.need<=0 || e.qty<=0) return;
      let hit=false;
      l.keys.forEach(k=>{ if(k===e.key || codeKeys_(e.key).indexOf(k)>=0) hit=true; });
      if(!hit) return;
      const take=Math.min(l.need, e.qty);
      l.need-=take; e.qty-=take;
    });
  });
}
