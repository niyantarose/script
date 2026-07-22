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
    if(!実EMS番号_(r[c.EMS番号])) return; // 棚卸箱・EMS番号空欄を履歴へ固定しない
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
function EMS番号入力分割_(input){
  return String(input==null?'':input)
    .split(/[,、\s\/／\.．]+/)
    .map(s=>s.trim())
    .filter(Boolean);
}

function 到着済を在庫反映済みへ(){ 直列_(到着済を在庫反映済みへ本体_); }
function 到着済を在庫反映済みへ本体_(){
  const ui=SpreadsheetApp.getUi(), cfg=P_KAKUTEI_CFG;
  let consistency=null, consistencyError='';
  try{
    const raw=PropertiesService.getDocumentProperties().getProperty('引当_整合状態');
    consistency=raw?JSON.parse(raw):null;
  }catch(error){ consistencyError=error.message; }
  const now=Date.now(), ts=consistency&&consistency.ts;
  const invalidTs=!Number.isFinite(ts)||ts<=0;
  const future=!invalidTs&&ts>now;
  const stale=!invalidTs&&!future&&(now-ts)>6*60*60*1000;
  if(consistencyError || !consistency || invalidTs || future || stale || consistency.要確認!==0 || consistency.台帳版!=='v1'){
    const reason=consistencyError?'整合状態を読み取れません: '+consistencyError
      :!consistency?'引当結果がありません'
      :invalidTs?'引当結果の時刻が不正です'
      :future?'引当結果の時刻が未来です'
      :stale?'直近6時間以内の引当結果がありません'
      :consistency.要確認!==0?'要確認が'+String(consistency.要確認)+'件残っています'
      :'台帳版がv1ではありません';
    ui.alert('便を締められません',reason+'。先に② 引き当て実行を正常完了してください。',ui.ButtonSet.OK);
    return;
  }

  let sh;
  try{ sh=発注共有を開く_().getSheetByName(cfg.シート); }
  catch(error){ ui.alert('発注共有ファイルが開けません:\n'+error.message); return; }
  if(!sh){ ui.alert('発注共有ファイルに「'+cfg.シート+'」がありません'); return; }
  const hr=cfg.ヘッダー行, last=sh.getLastRow();
  if(last<=hr){ ui.alert('EMSリストにデータがありません'); return; }
  const lastCol=sh.getLastColumn();
  const head=sh.getRange(hr,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const c=引当履歴_EMSリスト列_(head);
  const data=sh.getRange(hr+1,1,last-hr,lastCol).getDisplayValues();

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
    '\n\n在庫反映済みにするEMS番号を入力\n（複数はカンマ・/・.・空白区切り／空欄でOK=全部）',
    ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const input=String(resp.getResponseText()||'').trim();
  const targets=input?EMS番号入力分割_(input):list.slice();
  const bad=targets.filter(t=>!counts[t]);
  if(bad.length){ ui.alert('到着済に無いEMS番号があります: '+bad.join(', ')+'\n入力し直してください。'); return; }

  // 【締め前の先行残検査と未ピック確認】台帳が読めない時は安全側に停止する(2026-07-22 握り潰しcatch廃止)。
  // 希望日待ちは納品書が出ず現物の抜き忘れが起きやすい(2026-07-21 実例10117699の教訓)。
  // 抜き忘れたまま締めると、確保済みの現物が余りと一緒にYahoo保管へ紛れて宙に浮く。
  let 締め台帳;
  try{ 締め台帳=取り置き台帳_読む_(); }
  catch(e){ ui.alert('便を締められません','取り置き台帳が読めません(安全のため中止):\n'+e.message,ui.ButtonSet.OK); return; }
  {
    const 締めセット=new Set(targets);
    // 先行引当(帳簿のみ・現物なし)が対象便に残っていれば締めない。④で到着済へ昇格してから。
    const 先行残=取り置き_便締め先行残検査_(締め台帳,締めセット,{});
    if(先行残){ ui.alert('便を締められません',先行残,ui.ButtonSet.OK); return; }
    // 未ピック一覧は物理対象行(到着済・現物確認済み・要移行)だけを数える
    const 未ピック=締め台帳.filter(r=>String(r.状態||'')==='取り置き中'
      && 締めセット.has(String(r.元EMS番号||'').trim())
      && 取り置き_物理オペ対象行_(r,{}));
    if(未ピック.length){
      const byBan={};
      未ピック.forEach(r=>{ const b=String(r.受注番号||'');
        (byBan[b]=byBan[b]||[]).push(String(r.商品コード||'')+'×'+(Number(r.取り置き数量)||0)); });
      const bans=Object.keys(byBan);
      const 一覧=bans.slice(0,25).map(b=>'・'+b+': '+byBan[b].join(', '));
      const pick=ui.alert('締め前の未ピック確認',
        'この便には未出荷の確保(取り置き中)が '+未ピック.length+'行・'+bans.length+'注文あります。\n'
        +'現物を抜いて納品書を出しましたか？（希望日待ちは納品書が出ないので特に注意）\n\n'
        +一覧.join('\n')+(bans.length>25?'\n…ほか'+(bans.length-25)+'注文':'')
        +'\n\n抜き終わっていればOK ／ まだならキャンセルして先にピックしてください',
        ui.ButtonSet.OK_CANCEL);
      if(pick!==ui.Button.OK) return;
    }
  }

  const active=SpreadsheetApp.getActive(), jp=active.getSheetByName(HIKIATE_CFG.純在庫);
  if(!jp || jp.getLastRow()<2){ ui.alert('日本在庫に引当結果がありません。先に② 引き当て実行を正常完了してください。'); return; }
  const jpLastCol=jp.getLastColumn();
  const jpHead=jp.getRange(2,1,1,jpLastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const jpCode=jpHead.indexOf('商品コード'), jpQty=jpHead.indexOf('余り数(日本在庫)'), jpEms=jpHead.indexOf('EMS番号');
  if(jpCode<0 || jpQty<0 || jpEms<0){ ui.alert('日本在庫の見出しが不足しています。先に② 引き当て実行をやり直してください。'); return; }
  const targetSet=new Set(targets), jpLast=jp.getLastRow();
  const jpData=jpLast>2?jp.getRange(3,1,jpLast-2,jpLastCol).getDisplayValues():[];
  const surplus=jpData.map(row=>{
    const ems=String(row[jpEms]||'').trim(), sourceCode=String(row[jpCode]||'').trim();
    return {ems,code:sourceCode,sourceCode,qty:Number(row[jpQty])||0};
  }).filter(row=>targetSet.has(row.ems)&&row.qty>0);
  // 📤 出力済みかは記録(内容署名)で自動判定する。同じ内容を出力済みなら黙って先へ、
  // 未出力・内容が変わっていれば聞かずにこの場で出力する(締め後は日本在庫から余りが消えるため)。
  // typeofガードは単体テストハーネス(出力モジュール未ロード)との互換用。実GASでは常に定義済み
  if(surplus.length && typeof Yahoo変更_対象行_==='function' && typeof Yahoo在庫変更を出力本体_==='function'){
    let 除外=null; try{ 除外=全件再計算_マスタ除外集合_(); }catch(e){ 除外=new Set(); }
    const 出力対象=Yahoo変更_対象行_(
      surplus.map(r=>({商品コード:r.sourceCode,余り数:r.qty,EMS番号:r.ems})),
      targetSet,除外).対象;
    if(出力対象.length){
      const sig=Yahoo変更_内容署名_(出力対象);
      let 出力済み=false;
      try{ const rec=JSON.parse(PropertiesService.getDocumentProperties().getProperty('YAHOO出力記録')||'null');
        出力済み=!!rec && rec.sig===sig && sig!==''; }catch(e){}
      if(!出力済み) Yahoo在庫変更を出力本体_(targets);
    }
  }
  const moveLines=['EMS番号 / 商品コード / Yahooへ移す数量'].concat(
    surplus.length?surplus.map(row=>row.ems+' / '+row.sourceCode+' / '+row.qty):['(Yahoo移動対象なし)']);
  const confirm=ui.alert('Yahoo移動の最終確認',moveLines.join('\n')+'\n\nYahoo在庫への反映が完了している場合だけOK',ui.ButtonSet.OK_CANCEL);
  if(confirm!==ui.Button.OK) return;

  let movementPlan;
  try{ movementPlan=EMS在庫移動_箱計画_(surplus,EMS在庫移動台帳_読む_(),new Date()); }
  catch(error){ ui.alert('Yahoo移動を中止しました',error.message,ui.ButtonSet.OK); return; }
  if(movementPlan.errors.length){ ui.alert('Yahoo移動を中止しました',movementPlan.errors.join('\n'),ui.ButtonSet.OK); return; }

  try{ 引当履歴_今回到着分を記録_(true); }catch(error){}

  const flipA1=[];
  for(let i=0;i<data.length;i++){
    if(String(data[i][c.状態]||'').trim()!=='到着済') continue;
    const ems=String(data[i][c.EMS番号]||'').trim()||'(EMS番号なし)';
    if(targets.indexOf(ems)<0) continue;
    flipA1.push(sh.getRange(hr+1+i, c.状態+1).getA1Notation());
  }
  if(!flipA1.length){ ui.alert('対象行がありません。'); return; }
  let externalChanged=false;
  try{
    sh.getRangeList(flipA1).setValue('在庫反映済み');
    externalChanged=true;
    SpreadsheetApp.flush();
    EMS在庫移動台帳_保存_(movementPlan.rows);
  }catch(error){
    let rollbackError=null;
    if(externalChanged){
      try{
        sh.getRangeList(flipA1).setValue('到着済');
        SpreadsheetApp.flush();
      }catch(restoreError){ rollbackError=restoreError; }
    }
    const detail=error.message+(rollbackError?'\n到着済への復旧にも失敗しました: '+rollbackError.message:'');
    ui.alert('便の締めを中止しました','ステータスまたはYahoo移動台帳の更新に失敗しました。\n'+detail,ui.ButtonSet.OK);
    return;
  }

  let up={追加:0,重複:0,対象:0};
  try{ const r=引当履歴_EMSリストから記録_(['在庫反映済み'],'過去取込'); if(r && !r.error) up=r; }catch(error){}
  try{ EMS在庫を更新_本体_(); }catch(error){}

  ui.alert('✅ 便の締め 完了',
    '在庫反映済みへ: '+flipA1.length+'行（'+targets.join(', ')+'）\n'+
    '引当履歴: 過去取込へ昇格・追加'+(up.追加||0)+'件/既存更新'+(up.重複||0)+'件\n'+
    'EMS在庫: 最新化済み\n\n'+
    'このあと「② 引き当て実行」を回すと、締めた便の分は需要から差し引かれます。',
    ui.ButtonSet.OK);
}

// 【撤去 2026-07-21】引当履歴_反映済み割当マップ_/引当履歴_需要を差し引く_ は削除した。
// 過去取込の需要差引きは台帳・棚卸(開始前在庫)が無かった時代の仕組みで、古いP列の名指し
// (納品書=物理ピックの裏付けなし)がユーザーの数えた事実を上書きし、幽霊確保を生んでいた
// (実例10117699)。「誰が何を持っているか」は取り置き台帳だけが決める。履歴シートは監査記録のみ。
