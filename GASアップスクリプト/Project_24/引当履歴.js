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
  try{ src=発注共有を開く_().getSheetByName(cfg.シート); }
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

// ===== 便の締め: 到着済の箱を「在庫反映済み」にして、引当履歴を過去取込へ昇格 =====
// 使いどころ: その便の出荷作業が終わったとき。図形ボタンに「到着済を在庫反映済みへ」を割り当てて使う。
// やること: ①今のP列を履歴に記録(最新化) → ②指定EMS番号(空欄=到着済ぜんぶ)の行を在庫反映済みへ
//          → ③履歴を過去取込に昇格(以後、②やP列書き直しの需要差引きが効く) → ④EMS在庫を最新化
function 到着済を在庫反映済みへ(){ 直列_(到着済を在庫反映済みへ本体_); }
function 到着済を在庫反映済みへ本体_(){
  const ui=SpreadsheetApp.getUi(), cfg=P_KAKUTEI_CFG;
  // 【締めガード】⚠️が残ったまま締めると、実物を持たない注文のスタンプ(幽霊)が
  // 「過去便受取済み」として封印され、後から実在庫と合わなくなる(例: 10117220の6/22便)。
  // ②が保存した整合状態を確認し、⚠️あり/②未実行(6時間超)なら締める前に警告する。
  try{
    const raw=PropertiesService.getDocumentProperties().getProperty('引当_整合状態');
    const st=raw? JSON.parse(raw) : null;
    const 古い=!st || (Date.now()-(st.ts||0))>6*60*60*1000;
    if(古い || (st&&st.要確認>0)){
      const msg= 古い
        ? '直近6時間以内に ②引き当て実行 が走っていません。\n締める前に②を実行し、⚠️要確認が無いことを確認するのがおすすめです。'
        : '②の完了サマリに ⚠️要確認 '+st.要確認+'件 が残っています。\nこのまま締めると、実物を持たない注文のスタンプ(幽霊)が「過去便受取済み」として固定され、後から実在庫と合わなくなります。\n\n先に⚠️を解消（🔎商品診断／🔎整合チェック→🧹→②）してから締めるのがおすすめです。';
      const a=ui.alert('締めの前に確認', msg+'\n\nそれでも今すぐ締めますか？', ui.ButtonSet.OK_CANCEL);
      if(a!==ui.Button.OK) return;
    }
  }catch(e){ /* ガードが読めなくても締め自体は止めない */ }
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(e){ ui.alert('発注共有ファイルが開けません:\n'+e.message); return; }
  if(!sh){ ui.alert('発注共有ファイルに「'+cfg.シート+'」がありません'); return; }
  const hr=cfg.ヘッダー行, last=sh.getLastRow();
  if(last<=hr){ ui.alert('EMSリストにデータがありません'); return; }
  const lastCol=sh.getLastColumn();
  const head=sh.getRange(hr,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const c=引当履歴_EMSリスト列_(head);
  const data=sh.getRange(hr+1,1,last-hr,lastCol).getDisplayValues();

  // 今「到着済」の箱をEMS番号ごとに集計して選ばせる
  const counts={};
  data.forEach(r=>{
    if(String(r[c.状態]||'').trim()!=='到着済') return;
    const ems=String(r[c.EMS番号]||'').trim()||'(EMS番号なし)';
    counts[ems]=(counts[ems]||0)+1;
  });
  const list=Object.keys(counts);
  if(!list.length){ ui.alert('「到着済」の行がありません。'); return; }
  const resp=ui.prompt('到着済を在庫反映済みへ(便の締め)',
    '今「到着済」の箱:\n'+list.map(e=>'・'+e+'（'+counts[e]+'行）').join('\n')+
    '\n\n在庫反映済みにするEMS番号を入力\n（複数はカンマ区切り／空欄でOK=全部）',
    ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const input=String(resp.getResponseText()||'').trim();
  const targets= input? input.split(/[,、\s]+/).map(s=>s.trim()).filter(Boolean) : list.slice();
  const bad=targets.filter(t=>!counts[t]);
  if(bad.length){ ui.alert('到着済に無いEMS番号があります: '+bad.join(', ')+'\n入力し直してください。'); return; }

  // ① 締める前に、今のP列の割当を履歴へ記録(最新のP列を確実に残す)
  try{ 引当履歴_今回到着分を記録_(true); }catch(e){}

  // ② 対象行のステータスを「在庫反映済み」へ
  const flipA1=[];
  for(let i=0;i<data.length;i++){
    if(String(data[i][c.状態]||'').trim()!=='到着済') continue;
    const ems=String(data[i][c.EMS番号]||'').trim()||'(EMS番号なし)';
    if(targets.indexOf(ems)<0) continue;
    flipA1.push(sh.getRange(hr+1+i, c.状態+1).getA1Notation());
  }
  if(!flipA1.length){ ui.alert('対象行がありません。'); return; }
  sh.getRangeList(flipA1).setValue('在庫反映済み');
  SpreadsheetApp.flush();

  // ③ 履歴を過去取込へ昇格(既存の記録は取込区分/EMS状態が更新され、需要差引きの対象になる)
  let up={追加:0,重複:0,対象:0};
  try{ const r=引当履歴_EMSリストから記録_(['在庫反映済み'],'過去取込'); if(r && !r.error) up=r; }catch(e){}

  // ④ EMS在庫を最新化(反映済みがQUERYから消えるのを待つ)
  try{ EMS在庫を更新_本体_(); }catch(e){}

  ui.alert('✅ 便の締め 完了',
    '在庫反映済みへ: '+flipA1.length+'行（'+targets.join(', ')+'）\n'+
    '引当履歴: 過去取込へ昇格・追加'+(up.追加||0)+'件/既存更新'+(up.重複||0)+'件\n'+
    'EMS在庫: 最新化済み\n\n'+
    'このあと「② 引き当て実行」を回すと、締めた便の分は需要から差し引かれます。',
    ui.ButtonSet.OK);
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
